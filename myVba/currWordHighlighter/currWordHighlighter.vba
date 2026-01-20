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
Public previousSelectedText As String
Public isProcessingSelectionChange As Boolean ' 무한루프 방지 플래그

' 초기화 프로시저 (이 모듈이 로드될 때 호출)
Public Sub InitializeCurrWordHighlighter()
    isCurrWordHighlighterEnabled = True ' 초기 상태: 활성화
    previousSelectedText = ""
    isProcessingSelectionChange = False
End Sub

' 선택 영역 자동 검색 기능 토글
Public Sub ToggleCurrWordHighlighter()
    isCurrWordHighlighterEnabled = Not isCurrWordHighlighterEnabled
    
    If isCurrWordHighlighterEnabled Then
        ' 기능 활성화
        Call showMsg("선택 영역 자동 검색이 활성화되었습니다.", "알림", vbInformation, 1000)
    Else
        ' 기능 비활성화
        ' 이전 하이라이트 제거
        If previousSelectedText <> "" Then
            Call RemoveHighlight(previousSelectedText)
            previousSelectedText = ""
        End If
        Call showMsg("선택 영역 자동 검색이 비활성화되었습니다.", "알림", vbInformation, 1000)
    End If
End Sub

' 이전 하이라이트 제거 함수
Public Sub RemoveHighlight(searchText As String)
    On Error GoTo ErrorHandler
    
    Dim findRange As Range
    Dim originalRange As Range
    Dim scopeRange As Range
    
    ' 현재 커서 위치 저장 (Range 객체로 저장하여 Select 호출 방지)
    Set originalRange = Selection.Range.Duplicate
    
    ' 화면 업데이트 일시 중지 (이벤트 발생 감소)
    Application.ScreenUpdating = False
    
    ' 검색/제거 범위: 문서 전체
    Set scopeRange = ActiveDocument.Content
    
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
            ' findRange.HighlightColorIndex = wdNoHighlight
            findRange.Collapse wdCollapseEnd
        Loop
    End With
    
SafeExit:
    ' 화면 업데이트 재개
    Application.ScreenUpdating = True
    
    ' 원래 커서 위치로 복원 (플래그가 설정되어 있으면 Select 호출하지 않음)
    If Not isProcessingSelectionChange Then
        originalRange.Select
    Else
        ' 이벤트 처리 중이면 Range만 이동 (Select 호출 시 무한루프 발생)
        Selection.SetRange originalRange.Start, originalRange.End
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "하이라이트 제거 중 오류: " & Err.Description
    Application.ScreenUpdating = True
    ' 원래 커서 위치로 복원 시도
    On Error Resume Next
    If Not isProcessingSelectionChange Then
        originalRange.Select
    Else
        Selection.SetRange originalRange.Start, originalRange.End
    End If
End Sub

' 선택 영역이 변경될 때 호출되는 함수
' 이 함수는 clsAppEvents 클래스 모듈에서 appWord_WindowSelectionChange 이벤트로 호출됨
Public Sub HighlightCurrWord(ByVal targetRange As Range)
    If targetRange Is Nothing Then Exit Sub
    ' 기능이 비활성화되어 있으면 종료
    If Not isCurrWordHighlighterEnabled Then Exit Sub
    ' 무한루프 방지: 이미 처리 중이면 종료
    If isProcessingSelectionChange Then Exit Sub
    
    ' 처리 중 플래그 설정
    isProcessingSelectionChange = True
    
    On Error GoTo ErrorHandler
    
    Dim findRange As Range
    Dim originalRange As Range
    Dim scopeRange As Range
    Dim currentWord As String
    
    ' 텍스트가 선택(드래그)된 경우에는 하이라이트하지 않음
    If targetRange.Start <> targetRange.End Then
        isProcessingSelectionChange = False
        Exit Sub
    End If
    
    ' 커서만 있는 상태 -> 커서 위치 단어를 문서 전체에서 하이라이트
    currentWord = GetWordAtCursor(targetRange)
        
    ' 케이스 조건(camel/snake/pascal) 불만족 또는 빈 문자열이면 이전 하이라이트만 제거
    If currentWord = "" Or Not IsTargetIdentifierCase(currentWord) Then
        If previousSelectedText <> "" Then
            Call RemoveHighlight(previousSelectedText)
            previousSelectedText = ""
        End If
        isProcessingSelectionChange = False
        Exit Sub
    End If
        
    ' 이전 단어와 동일하면 유지
    If currentWord = previousSelectedText Then
        isProcessingSelectionChange = False
        Exit Sub
    End If
        
    ' 현재 커서 위치 저장 (Range 객체로 저장)
    Set originalRange = targetRange.Duplicate
        
    ' 하이라이트 범위: 문서 전체
    Set scopeRange = targetRange.Document.Content
        
    ' 이전 하이라이트 제거
    If previousSelectedText <> "" Then
        Call RemoveHighlight(previousSelectedText)
    End If
        
    ' 화면 업데이트 일시 중지 (이벤트 발생 감소)
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
                findRange.Shading.BackgroundPatternColor = GetTocHighlightColor()
            End If
            findRange.Collapse wdCollapseEnd
        Loop
    End With
    
    Application.ScreenUpdating = True
    
    previousSelectedText = currentWord
    isProcessingSelectionChange = False
    Exit Sub
    
ErrorHandler:
    Debug.Print "선택 영역 검색 및 하이라이트 적용 중 오류: " & Err.Description
    Application.ScreenUpdating = True
End Sub

' ======================
' 커서 단어 하이라이트용 유틸
' ======================

' 현재 커서 위치의 "단어"를 가져와 식별자 형태로 정리
Private Function GetWordAtCursor(ByVal targetRange As Range) As String
    On Error GoTo SafeExit
    
    Dim rng As Range
    Dim s As String
    
    If targetRange Is Nothing Then GoTo SafeExit
    Set rng = targetRange.Duplicate
    rng.Expand wdWord
    s = rng.Text
    
    ' 줄바꿈/탭 제거 및 트림
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Trim$(s)
    
    ' 앞/뒤 구두점 제거 (식별자용)
    s = TrimNonIdentifierEdges(s)
    
    GetWordAtCursor = s
    Exit Function
    
SafeExit:
    GetWordAtCursor = ""
End Function

' "밝은 녹색" 배경색 반환
Private Function GetTocHighlightColor() As Long
    ' 너무 진한 Lime(0,255,0) 대신 연한 녹색 계열 사용
    GetTocHighlightColor = RGB(198, 239, 206)
End Function

' camelCase / snake_case / PascalCase 중 하나인지 판별
Private Function IsTargetIdentifierCase(ByVal s As String) As Boolean
    If s = "" Then
        IsTargetIdentifierCase = False
        Exit Function
    End If
    
    IsTargetIdentifierCase = (IsCamelCaseIdentifier(s) Or IsPascalCaseIdentifier(s) Or IsSnakeCaseIdentifier(s))
End Function

Private Function IsCamelCaseIdentifier(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim hasUpper As Boolean
    
    ' underscore 금지
    If InStr(1, s, "_", vbBinaryCompare) > 0 Then Exit Function
    If Len(s) < 2 Then Exit Function
    
    ' 첫 글자 소문자
    ch = Mid$(s, 1, 1)
    If Not IsLowerAsciiLetter(ch) Then Exit Function
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If IsUpperAsciiLetter(ch) Then hasUpper = True
        If Not (IsLowerAsciiLetter(ch) Or IsUpperAsciiLetter(ch) Or IsAsciiDigit(ch)) Then Exit Function
    Next i
    
    IsCamelCaseIdentifier = hasUpper
End Function

Private Function IsPascalCaseIdentifier(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim hasLower As Boolean
    
    ' underscore 금지
    If InStr(1, s, "_", vbBinaryCompare) > 0 Then Exit Function
    If Len(s) < 2 Then Exit Function
    
    ' 첫 글자 대문자
    ch = Mid$(s, 1, 1)
    If Not IsUpperAsciiLetter(ch) Then Exit Function
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If IsLowerAsciiLetter(ch) Then hasLower = True
        If Not (IsLowerAsciiLetter(ch) Or IsUpperAsciiLetter(ch) Or IsAsciiDigit(ch)) Then Exit Function
    Next i
    
    IsPascalCaseIdentifier = hasLower
End Function

Private Function IsSnakeCaseIdentifier(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim hasUnderscore As Boolean
    
    If Len(s) < 3 Then Exit Function
    
    ' 반드시 underscore 포함
    hasUnderscore = (InStr(1, s, "_", vbBinaryCompare) > 0)
    If Not hasUnderscore Then Exit Function
    
    ' 처음/끝 underscore는 제외(엄격하게)
    If Left$(s, 1) = "_" Or Right$(s, 1) = "_" Then Exit Function
    ' 연속 underscore 제외(엄격하게)
    If InStr(1, s, "__", vbBinaryCompare) > 0 Then Exit Function
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If IsUpperAsciiLetter(ch) Then Exit Function ' snake는 대문자 없음
        If Not (IsLowerAsciiLetter(ch) Or IsAsciiDigit(ch) Or ch = "_") Then Exit Function
    Next i
    
    IsSnakeCaseIdentifier = True
End Function

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
        If IsIdentifierChar(beforeCh) Then GoTo SafeExit
    End If
    
    ' 오른쪽 문자
    If foundRange.End < scopeRange.End Then
        afterCh = doc.Range(foundRange.End, foundRange.End + 1).Text
        If IsIdentifierChar(afterCh) Then GoTo SafeExit
    End If
    
    IsBoundaryMatch = True
    Exit Function
    
SafeExit:
    IsBoundaryMatch = False
End Function

Private Function IsIdentifierChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsIdentifierChar = False
        Exit Function
    End If
    
    IsIdentifierChar = (IsLowerAsciiLetter(ch) Or IsUpperAsciiLetter(ch) Or IsAsciiDigit(ch) Or ch = "_")
End Function

Private Function IsLowerAsciiLetter(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsLowerAsciiLetter = (code >= 97 And code <= 122) ' a-z
End Function

Private Function IsUpperAsciiLetter(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsUpperAsciiLetter = (code >= 65 And code <= 90) ' A-Z
End Function

Private Function IsAsciiDigit(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsAsciiDigit = (code >= 48 And code <= 57) ' 0-9
End Function

' 문자열 양끝에서 식별자 문자(영문/숫자/_)가 아닌 것들을 제거
Private Function TrimNonIdentifierEdges(ByVal s As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim ch As String
    
    If s = "" Then
        TrimNonIdentifierEdges = ""
        Exit Function
    End If
    
    startPos = 1
    endPos = Len(s)
    
    Do While startPos <= endPos
        ch = Mid$(s, startPos, 1)
        If IsIdentifierChar(ch) Then Exit Do
        startPos = startPos + 1
    Loop
    
    Do While endPos >= startPos
        ch = Mid$(s, endPos, 1)
        If IsIdentifierChar(ch) Then Exit Do
        endPos = endPos - 1
    Loop
    
    If endPos < startPos Then
        TrimNonIdentifierEdges = ""
    Else
        TrimNonIdentifierEdges = Mid$(s, startPos, endPos - startPos + 1)
    End If
End Function

