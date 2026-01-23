Option Explicit

' =============================================================================
' Module: specFormatter
' Purpose:
'   - KIPO spec 문서에서 "【...】" 단락 규칙과 청구항 참조 관계를 이용해
'     Word 스타일(제목 1~9) 및 표/이미지 주변 문단 간격을 자동 정리합니다.
'
' Public API:
'   - formatKipoSpec()
' =============================================================================

' 실행 매크로.
' - "【...】" 단락에 제목 스타일(1~3)을 적용합니다.
' - 청구항 참조(제 n항/청구항 n)를 기반으로 제목 레벨을 추정해 조정합니다.
Public Sub formatKipoSpec()
    On Error GoTo ErrorHandler
    Call BeginCustomUndoRecord("formatKipoSpec")

    Dim doc As Document
    Set doc = ActiveDocument

    Call SetPageFormat(doc)
    Call SetDefaultParagraphFormat(doc)

    Call SetHeadingStyles(doc)
    Call ApplyHeadingStyles(doc)
    Call AdjustClaimHeadingLevels(doc)

    Call ShowNavigationPane()
    GoTo SafeExit

ErrorHandler:
    VBA.MsgBox "오류: " & Err.Description, vbCritical, "specFormatter"
    GoTo SafeExit

SafeExit:
    Call EndCustomUndoRecord()
End Sub

' 페이지 포맷 설정
Private Sub SetPageFormat(ByVal doc As Document)
    ' 페이지 여백을 설정합니다.
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(3)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.54)
        .RightMargin = CentimetersToPoints(2.54)
    End With
End Sub

' 기본 단락 포맷 설정
Private Sub SetDefaultParagraphFormat(ByVal doc As Document)
    With doc.Content.ParagraphFormat
        .Alignment = wdAlignParagraphJustify
        .LineSpacingRule = wdLineSpaceDouble
    End With
End Sub

' 제목(1-9) 스타일 설정
'  "【"가 있는 첫 번째 단락의 스타일로 설정합니다.
Private Sub SetHeadingStyles(ByVal doc As Document)
    Dim anchorPara As Paragraph
    Set anchorPara = FindFirstParagraphContaining(doc, "【")
    If anchorPara Is Nothing Then Exit Sub

    Dim anchorRng As Range
    Set anchorRng = anchorPara.Range.Duplicate

    Dim i As Long
    For i = 1 To 9
        With doc.Styles("제목 " & CStr(i))
            .AutomaticallyUpdate = False
            .Font = anchorRng.Font
            .ParagraphFormat = anchorRng.ParagraphFormat
            .ParagraphFormat.KeepWithNext = True

            If i = 1 Then
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
            Else
                .ParagraphFormat.Alignment = wdAlignParagraphJustify
            End If
        End With
    Next i
End Sub

Private Function FindFirstParagraphContaining( _
    ByVal doc As Document, _
    ByVal needle As String _
) As Paragraph
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If InStr(para.Range.Text, needle) > 0 Then
            Set FindFirstParagraphContaining = para
            Exit Function
        End If
    Next para

    Set FindFirstParagraphContaining = Nothing
End Function

Private Sub ApplyHeadingStyles(ByVal doc As Document)
    Dim rng As Range
    Set rng = doc.Content

    With rng.Find
        .Text = "【*"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
    End With

    Do While rng.Find.Execute
        Dim para As Paragraph
        Set para = rng.Paragraphs(1)
        Call ApplyHeadingStyleToBracketParagraph(para)

        rng.Collapse Direction:=wdCollapseEnd
        rng.MoveStart wdCharacter, 1
    Loop
End Sub

Private Sub ApplyHeadingStyleToBracketParagraph(ByVal para As Paragraph)
    Dim paraText As String
    paraText = para.Range.Text

    ' 제목 3
    If InStr(paraText, "【해결하고자 하는 과제") > 0 Or _
       InStr(paraText, "【기술적 과제") > 0 Or _
       InStr(paraText, "【과제의 해결 수단") > 0 Or _
       InStr(paraText, "【기술적 해결방법") > 0 Or _
       InStr(paraText, "【발명의 효과") > 0 Or _
       InStr(paraText, "【표") > 0 Or _
       InStr(paraText, "【수학식") > 0 Then
        para.Style = "제목 3"
        Exit Sub
    End If

    ' 제목 1
    If InStr(paraText, "【발명의 설명") > 0 Or _
       InStr(paraText, "【명세서") > 0 Or _
       InStr(paraText, "【청구범위") > 0 Or _
       InStr(paraText, "【청구의 범위") > 0 Or _
       InStr(paraText, "【요약서") > 0 Or _
       InStr(paraText, "【도면】") > 0 Then
        para.Style = "제목 1"
        Exit Sub
    End If

    ' 제목 2
    If InStr(paraText, "【") > 0 Then
        para.Style = "제목 2"
    End If
End Sub

Private Sub AdjustClaimHeadingLevels(ByVal doc As Document)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Call BuildClaimRangeIndex(doc, dict)
    If dict.Count = 0 Then Exit Sub

    Dim key As Variant
    For Each key In dict.Keys
        Call UpdateClaimHeadingLevel(doc, dict, CLng(key))
    Next key
End Sub

Private Sub BuildClaimRangeIndex(ByVal doc As Document, ByVal dict As Object)
    Dim rng As Range
    Set rng = doc.Content

    With rng.Find
        .Text = "【*"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
    End With

    Dim wasClaim As Boolean
    wasClaim = False

    Dim lastClaimNo As Long
    lastClaimNo = 0

    Do While rng.Find.Execute
        Dim para As Paragraph
        Set para = rng.Paragraphs(1)

        Dim paraText As String
        paraText = para.Range.Text

        If wasClaim Then
            Call SetClaimEnd(dict, lastClaimNo, para.Range.Start - 1)
            wasClaim = False
        End If

        Dim claimNo As Long
        If TryExtractClaimNumber(paraText, claimNo) Then
            Dim nextParaStart As Long
            nextParaStart = para.Range.End

            Dim paraIndex As Long
            paraIndex = doc.Range(0, para.Range.Start).Paragraphs.Count + 1

            dict.Add claimNo, Array(2, nextParaStart, -1, paraIndex)
            wasClaim = True
            lastClaimNo = claimNo
        End If

        rng.Collapse Direction:=wdCollapseEnd
        rng.MoveStart wdCharacter, 1
    Loop

    If wasClaim Then
        Call SetClaimEnd( _
            dict, _
            lastClaimNo, _
            doc.Paragraphs(doc.Paragraphs.Count).Range.End _
        )
    End If
End Sub

Private Sub SetClaimEnd(ByVal dict As Object, ByVal claimNo As Long, ByVal endPos As Long)
    If Not dict.Exists(claimNo) Then Exit Sub

    Dim info As Variant
    info = dict(claimNo)
    info(2) = endPos
    dict(claimNo) = info
End Sub

Private Function TryExtractClaimNumber(ByVal text As String, ByRef outNo As Long) As Boolean
    On Error GoTo SafeExit

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "청구항[ ]?([0-9]+)"
    re.Global = False
    re.IgnoreCase = False

    If Not re.Test(text) Then GoTo SafeExit

    Dim matches As Object
    Set matches = re.Execute(text)

    outNo = CLng(matches(0).SubMatches(0))
    TryExtractClaimNumber = True
    Exit Function

SafeExit:
    outNo = 0
    TryExtractClaimNumber = False
End Function

Private Sub UpdateClaimHeadingLevel( _
    ByVal doc As Document, _
    ByVal dict As Object, _
    ByVal claimNo As Long _
)
    Dim info As Variant
    info = dict(claimNo)

    Dim startPos As Long
    startPos = CLng(info(1))

    Dim endPos As Long
    endPos = CLng(info(2))
    If endPos < startPos Then Exit Sub

    Dim paraRng As Range
    Set paraRng = doc.Range(Start:=startPos, End:=endPos)

    Dim upperClaimNo As Long
    upperClaimNo = FindSmallestClaimReferenceNo(paraRng)

    If upperClaimNo <= 0 Then Exit Sub
    If Not dict.Exists(upperClaimNo) Then Exit Sub

    Dim newLevel As Long
    newLevel = CLng(dict(upperClaimNo)(0)) + 1

    Dim claimParaIdx As Long
    claimParaIdx = CLng(info(3))

    Dim claimRng As Range
    Set claimRng = doc.Paragraphs(claimParaIdx).Range
    claimRng.Style = "제목 " & CStr(newLevel)

    info(0) = newLevel
    dict(claimNo) = info
End Sub

Private Function FindSmallestClaimReferenceNo(ByVal rng As Range) As Long
    On Error GoTo SafeExit

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "제\s*\d+항|청구항\s*\d+"

    Dim matches As Object
    Set matches = re.Execute(rng.Text)

    Dim smallest As Long
    smallest = 0

    Dim m As Object
    For Each m In matches
        Dim n As Long
        n = ExtractFirstLong(CStr(m.Value))
        If n <= 0 Then GoTo ContinueLoop

        If smallest = 0 Or n < smallest Then
            smallest = n
        End If

ContinueLoop:
    Next m

    FindSmallestClaimReferenceNo = smallest
    Exit Function

SafeExit:
    FindSmallestClaimReferenceNo = 0
End Function

Private Function ExtractFirstLong(ByVal s As String) As Long
    On Error GoTo SafeExit

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(\d+)"

    If Not re.Test(s) Then GoTo SafeExit

    Dim matches As Object
    Set matches = re.Execute(s)
    ExtractFirstLong = CLng(matches(0).SubMatches(0))
    Exit Function

SafeExit:
    ExtractFirstLong = 0
End Function


Private Sub ShowNavigationPane()
    On Error Resume Next
    With CommandBars("Navigation")
        .Visible = True
    End With
    On Error GoTo 0
End Sub
