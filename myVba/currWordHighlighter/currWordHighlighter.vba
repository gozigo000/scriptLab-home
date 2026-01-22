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

' 이전 하이라이트 정보를 담는 타입
Public Type TPrevHighlightInfo
    Word As String
    ScopeStart As Long ' 직전에 하이라이트를 적용한 "섹션" 범위 시작
    ScopeEnd As Long   ' 직전에 하이라이트를 적용한 "섹션" 범위 끝(End, exclusive)
End Type

' 모듈 레벨 변수
Public isCurrWordHighlighterEnabled As Boolean ' 기능 ON/OFF 토글 (True=활성화)
Public isProcessingSelectionChange As Boolean ' 무한루프 방지 플래그
Public gPrevHighlight As TPrevHighlightInfo

' (MARK) 초기화
Public Sub InitializeCurrWordHighlighter()
    isCurrWordHighlighterEnabled = False ' 초기 상태: 비활성화
    isProcessingSelectionChange = False
    gPrevHighlight.Word = ""
    gPrevHighlight.ScopeStart = 0
    gPrevHighlight.ScopeEnd = 0
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
        If gPrevHighlight.Word <> "" Then
            Call RemoveHighlight(ActiveDocument, gPrevHighlight.Word, GetPreviousHighlightedScopeRange(ActiveDocument))
            gPrevHighlight.Word = ""
            gPrevHighlight.ScopeStart = 0
            gPrevHighlight.ScopeEnd = 0
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
    
    ' 현재 섹션(헤딩 단위) 범위로 검색/하이라이트 범위를 제한
    Dim scopeRange As Range
    Set scopeRange = GetCurrentHeadingRange(targetRange)
    If scopeRange Is Nothing Then
        ' 헤딩이 없는 문서/구간이면 기존처럼 문서 전체를 하나의 섹션으로 간주
        Set scopeRange = doc.Content
    End If
    
    Dim varRng As Range
    Set varRng = GetVariableRangeAtPos(doc, targetRange.Start)
    
    Dim currentWord As String
    currentWord = GetRangeText(varRng)

    ' 이전 단어와 동일하면 유지(단, "같은 섹션"일 때만)
    If currentWord = gPrevHighlight.Word Then
        If gPrevHighlight.ScopeStart = scopeRange.Start And _
            gPrevHighlight.ScopeEnd = scopeRange.End _
        Then
            GoTo Cleanup
        End If
        ' 단어는 같지만 섹션이 바뀐 경우:
        ' - 이전 섹션의 하이라이트는 지우고
        ' - 현재 섹션에 다시 적용해야 함 (아래 로직으로 진행)
    End If

    ' variableCase.vba 유틸: 케이스 판별 (camel/snake/pascal)
    If currentWord = "" Or Not ( _
        IsCamelCase(currentWord) Or _
        IsPascalCase(currentWord) Or _
        IsSnakeCase(currentWord) Or _
        IsScreamingSnakeCase(currentWord) _
    ) Then
        If gPrevHighlight.Word <> "" Then
            Call BeginCustomUndoRecord()
            Call RemoveHighlight(doc, gPrevHighlight.Word, GetPreviousHighlightedScopeRange(doc))
            gPrevHighlight.Word = ""
            gPrevHighlight.ScopeStart = 0
            gPrevHighlight.ScopeEnd = 0
        End If
        GoTo Cleanup
    End If
    
    
    Dim findRange As Range
    
    ' 이전 하이라이트 제거
    If gPrevHighlight.Word <> "" Then
        Call BeginCustomUndoRecord()
        Call RemoveHighlight(doc, gPrevHighlight.Word, GetPreviousHighlightedScopeRange(doc))
    End If
    
    Application.ScreenUpdating = False
    
    ' 현재 섹션에 하이라이트 적용도 동일 UndoRecord로 포함
    Call BeginCustomUndoRecord()
    
    ' 현재 섹션 범위에서 동일 단어 검색
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
            
            ' 다음 검색 시작점으로 이동하되, 검색 범위 End는 "현재 섹션"으로 고정
            Dim nextPos As Long
            nextPos = findRange.End
            If nextPos >= scopeRange.End Then Exit Do
            
            findRange.Start = nextPos
            findRange.End = scopeRange.End
        Loop
    End With
    
    gPrevHighlight.Word = currentWord
    gPrevHighlight.ScopeStart = scopeRange.Start
    gPrevHighlight.ScopeEnd = scopeRange.End
    
Cleanup:
    On Error Resume Next
    Call EndCustomUndoRecord
    Application.ScreenUpdating = True
    isProcessingSelectionChange = False
    Exit Sub
    
ErrorHandler:
    Debug.Print "선택 영역 검색 및 하이라이트 적용 중 오류: " & Err.Description
    Resume Cleanup
End Sub

' (MARK) 이전 하이라이트 제거
Public Sub RemoveHighlight( _
    ByVal doc As Document, _
    ByVal searchText As String, _
    Optional ByVal scopeRange As Range _
)
    On Error GoTo ErrorHandler
    
    If doc Is Nothing Then Exit Sub
    If searchText = "" Then Exit Sub
    
    Dim findRange As Range
    
    Application.ScreenUpdating = False
    
    ' 검색/제거 범위: 기본은 문서 전체(호출자가 범위를 주면 그 범위만)
    If scopeRange Is Nothing Then
        Set scopeRange = doc.Content
    End If
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
            
            ' 다음 검색 시작점으로 이동하되, 검색 범위 End는 scopeRange로 고정
            Dim nextPos As Long
            nextPos = findRange.End
            If nextPos >= scopeRange.End Then Exit Do
            
            findRange.Start = nextPos
            findRange.End = scopeRange.End
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

' (MARK) 직전에 하이라이트한 섹션 범위를 복원 (없으면 doc.Content)
Private Function GetPreviousHighlightedScopeRange(ByVal doc As Document) As Range
    On Error GoTo Fallback
    
    If doc Is Nothing Then GoTo Fallback
    If gPrevHighlight.ScopeStart <= 0 Then GoTo Fallback
    If gPrevHighlight.ScopeEnd <= 0 Then GoTo Fallback
    If gPrevHighlight.ScopeEnd < gPrevHighlight.ScopeStart Then GoTo Fallback
    
    Dim s As Long
    Dim e As Long
    s = gPrevHighlight.ScopeStart
    e = gPrevHighlight.ScopeEnd
    
    ' 문서 편집 등으로 값이 어긋났을 수 있어 클램프
    If s < doc.Content.Start Then s = doc.Content.Start
    If e > doc.Content.End Then e = doc.Content.End
    If e < s Then GoTo Fallback
    
    Set GetPreviousHighlightedScopeRange = doc.Range(s, e)
    Exit Function
    
Fallback:
    If doc Is Nothing Then
        Set GetPreviousHighlightedScopeRange = Nothing
    Else
        Set GetPreviousHighlightedScopeRange = doc.Content
    End If
End Function

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
