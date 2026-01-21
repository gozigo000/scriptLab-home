Option Explicit

' ============================================================
' ImportMyVbaToNormal.bas
' - 지정 폴더의 .bas/.cls/.frm 파일을 Normal.dotm에 일괄 반영
' - 동일 모듈명(VB_Name) 존재 시: 삭제(Remove) 후 Import
' - ThisDocument는 삭제 불가 → 코드 내용을 파일 내용으로 교체
'
' 사전 조건:
' - Word 보안 설정에서 "VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스" 활성화
' ============================================================

' TODO: 본인 PC의 'scriptLab-home' 폴더 경로로 수정하세요.
Private Const rootDir As String = "C:\Users\gozig\Desktop\scriptLab-home"

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Public Sub ImportMyVbaToNormal()
    ' 변환 스크립트 실행
    If Not RunConvertMyVbaToDist() Then Exit Sub

    Const folderMsWordObject As String = rootDir & "\dist\MsWord객체\"
    Const folderModules As String = rootDir & "\dist\모듈\"
    Const folderClassModules As String = rootDir & "\dist\클래스모듈\"
    Const folderForms As String = rootDir & "\dist\폼\"

    If Not FolderExists(folderMsWordObject) Then
        VBA.MsgBox "폴더를 찾을 수 없습니다: " & folderMsWordObject, vbCritical
        Exit Sub
    End If
    If Not FolderExists(folderModules) Then
        VBA.MsgBox "폴더를 찾을 수 없습니다: " & folderModules, vbCritical
        Exit Sub
    End If
    If Not FolderExists(folderClassModules) Then
        VBA.MsgBox "폴더를 찾을 수 없습니다: " & folderClassModules, vbCritical
        Exit Sub
    End If
    If Not FolderExists(folderForms) Then
        VBA.MsgBox "폴더를 찾을 수 없습니다: " & folderForms, vbCritical
        Exit Sub
    End If

    Dim vbProj As Object
    Set vbProj = NormalTemplate.VBProject

    Dim importedCount As Long
    importedCount = 0

    ' MsWord객체(예: ThisDocument) → 코드 교체 방식이 적용될 수 있음
    importedCount = importedCount + ImportByPattern(vbProj, folderMsWordObject, "*.cls")
    ' 표준 모듈
    importedCount = importedCount + ImportByPattern(vbProj, folderModules, "*.bas")
    ' 클래스 모듈
    importedCount = importedCount + ImportByPattern(vbProj, folderClassModules, "*.cls")
    ' 폼
    importedCount = importedCount + ImportByPattern(vbProj, folderForms, "*.frm")

    NormalTemplate.Save

    VBA.MsgBox "완료: " & importedCount & "개 파일을 Normal.dotm에 반영했습니다.", vbInformation
End Sub

Private Function RunConvertMyVbaToDist() As Boolean
    On Error GoTo EH

    Dim scriptPath As String
    scriptPath = rootDir & "\distor\ConvertMyVbaToDist.ps1"

    If Len(Dir$(scriptPath)) = 0 Then
        VBA.MsgBox "PowerShell 스크립트를 찾을 수 없습니다: " & scriptPath, vbCritical
        RunConvertMyVbaToDist = False
        Exit Function
    End If

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim oldDir As String
    oldDir = sh.CurrentDirectory
    sh.CurrentDirectory = rootDir

    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & QuoteArg(scriptPath)

    Dim exitCode As Long
    exitCode = CLng(sh.Run(cmd, 0, True)) ' 0=hidden, True=wait

    sh.CurrentDirectory = oldDir

    If exitCode <> 0 Then
        VBA.MsgBox "변환 스크립트 실행 실패 (exit code=" & exitCode & "):" & vbCrLf & scriptPath, vbCritical
        RunConvertMyVbaToDist = False
        Exit Function
    End If

    RunConvertMyVbaToDist = True
    Exit Function

EH:
    On Error Resume Next
    If Not sh Is Nothing Then sh.CurrentDirectory = oldDir
    VBA.MsgBox "변환 스크립트 실행 중 오류: " & Err.Description, vbCritical
    RunConvertMyVbaToDist = False
End Function

Private Function QuoteArg(ByVal s As String) As String
    QuoteArg = """" & Replace$(s, """", """""") & """"
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    FolderExists = (Len(Dir$(folderPath, vbDirectory)) > 0)
    Exit Function
EH:
    FolderExists = False
End Function

Private Function ImportByPattern(ByVal vbProj As Object, ByVal folder As String, ByVal pattern As String) As Long
    Dim f As String
    Dim fullPath As String
    Dim count As Long
    count = 0

    f = Dir$(folder & pattern)
    Do While Len(f) > 0
        fullPath = folder & f
        If ImportOne(vbProj, fullPath) Then
            count = count + 1
        End If
        f = Dir$()
    Loop

    ImportByPattern = count
End Function

Private Function ImportOne(ByVal vbProj As Object, ByVal filePath As String) As Boolean
    On Error GoTo EH

    Dim vbName As String
    vbName = GetVBNameFromExportFile(filePath)
    If Len(vbName) = 0 Then
        vbName = GetBaseName(filePath)
    End If

    ' ThisDocument는 제거할 수 없으므로 코드 교체 방식으로 처리
    If StrComp(vbName, "ThisDocument", vbTextCompare) = 0 Then
        ReplaceDocumentModuleCode vbProj, "ThisDocument", ReadAllText(filePath)
        ImportOne = True
        Exit Function
    End If

    ' 동일 이름의 컴포넌트가 있으면 삭제
    Dim comp As Object
    Set comp = GetVBComponentByName(vbProj, vbName)
    If Not comp Is Nothing Then
        If CLng(CallByName(comp, "Type", VbGet)) <> vbext_ct_Document Then
            vbProj.VBComponents.Remove comp
        End If
    End If

    vbProj.VBComponents.Import filePath

    ImportOne = True
    Exit Function

EH:
    ImportOne = False
End Function

Private Function GetVBComponentByName(ByVal vbProj As Object, ByVal name As String) As Object
    On Error GoTo NotFound
    Dim c As Object
    Set c = vbProj.VBComponents(name)
    Set GetVBComponentByName = c
    Exit Function
NotFound:
    Set GetVBComponentByName = Nothing
End Function

Private Sub ReplaceDocumentModuleCode(ByVal vbProj As Object, ByVal docModuleName As String, ByVal newCode As String)
    On Error GoTo EH

    Dim comp As Object
    Set comp = vbProj.VBComponents(docModuleName)

    Dim codeMod As Object
    Set codeMod = CallByName(comp, "CodeModule", VbGet)

    Dim totalLines As Long
    totalLines = CLng(CallByName(codeMod, "CountOfLines", VbGet))
    If totalLines > 0 Then
        CallByName codeMod, "DeleteLines", VbMethod, 1, totalLines
    End If

    If Len(newCode) > 0 Then
        CallByName codeMod, "AddFromString", VbMethod, newCode
    End If
    Exit Sub

EH:
    ' 무시 (보안 설정/권한 문제 등)
End Sub

Private Function GetVBNameFromExportFile(ByVal filePath As String) As String
    On Error GoTo EH

    Dim s As String
    s = ReadAllText(filePath)

    ' Attribute VB_Name = "ModuleName"
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\s*Attribute\s+VB_Name\s*=\s*""([^""]+)"""
    re.Global = False
    re.IgnoreCase = True
    re.Multiline = True

    If re.Test(s) Then
        Dim m As Object
        Set m = re.Execute(s)(0)
        GetVBNameFromExportFile = CStr(m.SubMatches(0))
    Else
        GetVBNameFromExportFile = vbNullString
    End If

    Exit Function
EH:
    GetVBNameFromExportFile = vbNullString
End Function

Private Function GetBaseName(ByVal filePath As String) As String
    Dim p As Long
    Dim f As String
    f = filePath

    p = InStrRev(f, "\")
    If p > 0 Then f = Mid$(f, p + 1)

    p = InStrRev(f, ".")
    If p > 0 Then f = Left$(f, p - 1)

    GetBaseName = f
End Function

Private Function ReadAllText(ByVal filePath As String) As String
    On Error GoTo EH

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadAllText = stm.ReadText(-1)
    stm.Close
    Exit Function

EH:
    ReadAllText = vbNullString
End Function

