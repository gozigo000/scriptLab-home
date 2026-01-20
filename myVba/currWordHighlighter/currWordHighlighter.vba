Option Explicit

' (MARK) 선택 영역 자동 검색 섹션
' ----------------------
' 페이지 로드 시 이벤트 핸들러 등록 (기본적으로 비활성화 상태)
' 사용자가 버튼을 클릭하면 활성화됨
'
' 사용 방법:
' 1. 이 모듈을 Word VBA 프로젝트에 추가합니다.
' 2. MsgBox 모듈을 Word VBA 프로젝트에 추가합니다.
'    (MsgBox.vba 파일 참조)
' 3. clsAppEvents 클래스 모듈을 Word VBA 프로젝트에 추가합니다.
'    (clsAppEvents.cls 파일 참조)
' 4. ThisDocument 모듈에 다음 코드를 추가합니다:
'
'    Dim myAppEvents As clsAppEvents
'
'    Private Sub Document_Open()
'        Call InitializeCurrWordHighlighter
'        Set myAppEvents = New clsAppEvents
'        Set myAppEvents.appWord = Word.Application
'    End Sub
'
' 5. 기능을 활성화하려면 ToggleCurrWordHighlighter 서브루틴을 호출합니다.
'    예: 매크로 버튼이나 단축키에 ToggleCurrWordHighlighter 할당

' 모듈 레벨 변수
Public isCurrWordHighlighterEnabled As Boolean ' 기능 ON/OFF 토글 (True=활성화)
Public previousHighlightedWord As String
Public isProcessingSelectionChange As Boolean ' 무한루프 방지 플래그

' (MARK) 초기화
Public Sub InitializeCurrWordHighlighter()
    isCurrWordHighlighterEnabled = True ' 초기 상태: 활성화
    previousHighlightedWord = ""
    isProcessingSelectionChange = False
End Sub

' (MARK) 단어 하이라이트 기능 토글
Public Sub ToggleCurrWordHighlighter()
    isCurrWordHighlighterEnabled = Not isCurrWordHighlighterEnabled
    
    If isCurrWordHighlighterEnabled Then
        ' 기능 활성화
        Call showMsg("단어 하이라이트 기능이 활성화되었습니다.", "알림", vbInformation, 1000)
    Else
        ' 기능 비활성화
        ' 이전 하이라이트 제거
        If previousHighlightedWord <> "" Then
            Call RemoveHighlight(ActiveDocument, previousHighlightedWord)
            previousHighlightedWord = ""
        End If
        Call showMsg("단어 하이라이트 기능이 비활성화되었습니다.", "알림", vbInformation, 1000)
    End If
End Sub

' (MARK) 단어 하이라이트
Public Sub HighlightCurrWord(ByVal targetRange As Range)
    If targetRange Is Nothing Then Exit Sub
    ' 기능이 비활성화되어 있으면 종료
    If Not isCurrWordHighlighterEnabled Then Exit Sub
    ' 텍스트가 선택(드래그)된 경우에는 하이라이트하지 않음
    If targetRange.Start <> targetRange.End Then Exit Sub
    ' 무한루프 방지: 이미 처리 중이면 종료
    If isProcessingSelectionChange Then Exit Sub
    
    ' 처리 중 플래그 설정
    isProcessingSelectionChange = True
    
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Set doc = targetRange.Document
    
    Dim varRng As Range
    Set varRng = GetVariableRangeAtPos(doc, targetRange.Start)
    
    Dim currentWord As String
    currentWord = GetRangeText(varRng)

    ' 이전 단어와 동일하면 유지
    If currentWord = previousHighlightedWord Then GoTo Cleanup

    ' variableCase.vba 유틸: 케이스 판별 (camel/snake/pascal)
    If currentWord = "" Or Not ( _
        IsCamelCase(currentWord) Or _
        IsPascalCase(currentWord) Or _
        IsSnakeCase(currentWord) Or _
        IsScreamingSnakeCase(currentWord) _
    ) Then
        If previousHighlightedWord <> "" Then
            Call RemoveHighlight(doc, previousHighlightedWord)
            previousHighlightedWord = ""
        End If
        GoTo Cleanup
    End If
    
    
    Dim findRange As Range
    Dim scopeRange As Range
    Set scopeRange = doc.Content
    
    ' 이전 하이라이트 제거
    If previousHighlightedWord <> "" Then
        Call RemoveHighlight(doc, previousHighlightedWord)
    End If
    
    Application.ScreenUpdating = False
    
    ' 문서 전체에서 동일 단어 검색
    Set findRange = scopeRange.Duplicate
    With findRange.Find
        .ClearFormatting
        .Text = currentWord
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            If IsBoundaryMatch(findRange, scopeRange) Then
                ' 연한 녹색 배경
                findRange.Shading.BackgroundPatternColor = RGB(198, 239, 206)
            End If
            findRange.Collapse wdCollapseEnd
        Loop
    End With
    
    previousHighlightedWord = currentWord
    
Cleanup:
    On Error Resume Next
    Application.ScreenUpdating = True
    isProcessingSelectionChange = False
    Exit Sub
    
ErrorHandler:
    Debug.Print "선택 영역 검색 및 하이라이트 적용 중 오류: " & Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    isProcessingSelectionChange = False
End Sub

' (MARK) 이전 하이라이트 제거
Public Sub RemoveHighlight(ByVal doc As Document, ByVal searchText As String)
    On Error GoTo ErrorHandler
    
    If doc Is Nothing Then Exit Sub
    If searchText = "" Then Exit Sub
    
    Dim findRange As Range
    Dim scopeRange As Range
    
    Application.ScreenUpdating = False
    
    ' 검색/제거 범위: 문서 전체
    Set scopeRange = doc.Content
    Set findRange = scopeRange.Duplicate
    
    With findRange.Find
        .ClearFormatting
        .Text = searchText
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        
        ' 모든 일치 항목 찾아서 하이라이트 제거
        Do While .Execute
            If IsBoundaryMatch(findRange, scopeRange) Then
                findRange.Shading.BackgroundPatternColor = wdColorAutomatic ' 자동 색상
            End If
            findRange.Collapse wdCollapseEnd
        Loop
    End With
    
Cleanup:
    On Error Resume Next
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Debug.Print "하이라이트 제거 중 오류: " & Err.Description
    Resume Cleanup
End Sub

' ======================
' 커서 단어 하이라이트용 유틸
' ======================

' Find로 잡힌 구간이 "부분 문자열"이 아닌지(좌/우 경계가 식별자 문자가 아닌지) 확인
Private Function IsBoundaryMatch(ByVal foundRange As Range, ByVal scopeRange As Range) As Boolean
    On Error GoTo SafeExit
    
    Dim beforeCh As String
    Dim afterCh As String
    Dim doc As Document
    
    Set doc = scopeRange.Document
    
    ' 왼쪽 문자
    If foundRange.Start > scopeRange.Start Then
        beforeCh = doc.Range(foundRange.Start - 1, foundRange.Start).Text
        If IsVariableChar(beforeCh) Then GoTo SafeExit
    End If
    
    ' 오른쪽 문자
    If foundRange.End < scopeRange.End Then
        afterCh = doc.Range(foundRange.End, foundRange.End + 1).Text
        If IsVariableChar(afterCh) Then GoTo SafeExit
    End If
    
    IsBoundaryMatch = True
    Exit Function
    
SafeExit:
    IsBoundaryMatch = False
End Function
