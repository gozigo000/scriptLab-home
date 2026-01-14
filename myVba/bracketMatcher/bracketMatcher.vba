' (MARK) 괄호 매칭 및 하이라이트 섹션
' ----------------------
' 커서가 괄호 옆에 있을 때 해당 괄호의 짝을 찾아 하이라이트하는 기능
'
' 사용 방법:
' 1. 이 모듈을 Word VBA 프로젝트에 추가합니다.
' 2. clsAppEvents 클래스 모듈을 수정하여 OnBracketMatch 함수를 호출하도록 합니다.
'    (clsAppEvents.cls 파일의 appWord_WindowSelectionChange 이벤트에 추가)
' 3. ThisDocument 모듈에서 Document_Open 이벤트에 InitializeBracketMatcher 호출 추가

' 모듈 레벨 변수
Public previousBracketRanges As Collection ' 이전에 하이라이트된 괄호 범위들 저장
Public previousBracketColors As Collection ' 이전 괄호의 원래 배경색 저장
Public isProcessingBracketMatch As Boolean ' 무한루프 방지 플래그
Public isUndoRecordActive As Boolean ' UndoRecord가 활성화되어 있는지 추적

' 초기화 프로시저
Public Sub InitializeBracketMatcher()
    Set previousBracketRanges = New Collection
    Set previousBracketColors = New Collection
    isProcessingBracketMatch = False
    isUndoRecordActive = False
End Sub

' 괄호 매칭 및 하이라이트 함수
' 이 함수는 clsAppEvents 클래스 모듈에서 appWord_WindowSelectionChange 이벤트로 호출됨
Public Sub OnBracketMatch()
    ' 선택 영역의 길이가 0이 아니면 종료 (텍스트가 선택된 경우)
    If Selection.Type <> wdSelectionIP Then
        ' 선택 영역이 있으면 이전 하이라이트만 제거하고 종료
        Call RemoveBracketHighlight
        ' 하이라이트가 완전히 제거되었으므로 CustomRecord 종료
        If isUndoRecordActive Then
            On Error Resume Next
            Application.UndoRecord.EndCustomRecord
            On Error Resume Next
            isUndoRecordActive = False
        End If
        Exit Sub
    End If
    
    ' 무한루프 방지: 이미 처리 중이면 종료
    If isProcessingBracketMatch Then
        Exit Sub
    End If
    
    ' 처리 중 플래그 설정
    isProcessingBracketMatch = True
    
    On Error GoTo ErrorHandler
    
    Dim currentChar As String
    Dim cursorPos As Long
    Dim docRange As Range
    Dim bracketRange As Range
    Dim matchedRange As Range
    Dim originalRange As Range
    Dim charBefore As String
    Dim charAfter As String
    Dim rangeBefore As Range
    Dim rangeAfter As Range
    Dim hasBracketBefore As Boolean
    Dim hasBracketAfter As Boolean
    
    ' 현재 커서 위치 저장
    Set originalRange = Selection.Range.Duplicate
    cursorPos = Selection.Start
    
    ' 이전 하이라이트 제거
    Call RemoveBracketHighlight
    
    ' 커서 앞 문자 확인
    hasBracketBefore = False
    If cursorPos > 0 Then
        Set rangeBefore = ActiveDocument.Range(cursorPos - 1, cursorPos)
        charBefore = rangeBefore.Text
        hasBracketBefore = IsBracket(charBefore)
    End If
    
    ' 커서 뒤 문자 확인
    hasBracketAfter = False
    If cursorPos < ActiveDocument.Content.End Then
        Set rangeAfter = ActiveDocument.Range(cursorPos, cursorPos + 1)
        charAfter = rangeAfter.Text
        hasBracketAfter = IsBracket(charAfter)
    End If
    
    ' 괄호가 없으면 CustomRecord 종료
    If Not hasBracketBefore And Not hasBracketAfter Then
        If isUndoRecordActive Then
            On Error Resume Next
            Application.UndoRecord.EndCustomRecord
            On Error GoTo ErrorHandler
            isUndoRecordActive = False
        End If
        GoTo RestoreCursor
    End If
    
    ' 커서 앞뒤에 같은 종류의 여는 괄호가 연속으로 있는 경우, 바깥쪽(더 앞쪽) 괄호 선택
    If hasBracketBefore And hasBracketAfter Then
        If IsOpenBracket(charBefore) And IsOpenBracket(charAfter) Then
            ' 바깥쪽 괄호(앞쪽) 선택
            Set bracketRange = rangeBefore.Duplicate
            Set matchedRange = FindMatchingBracket(bracketRange, charBefore)
            
            If Not matchedRange Is Nothing Then
                Call HighlightBracketPair(bracketRange, matchedRange)
                GoTo RestoreCursor
            End If
        ' 커서 앞뒤에 같은 종류의 닫는 괄호가 연속으로 있는 경우, 바깥쪽(더 뒤쪽) 괄호 선택
        ElseIf IsCloseBracket(charBefore) And IsCloseBracket(charAfter) Then
            ' 바깥쪽 괄호(뒤쪽) 선택
            Set bracketRange = rangeAfter.Duplicate
            Set matchedRange = FindMatchingBracket(bracketRange, charAfter)
            
            If Not matchedRange Is Nothing Then
                Call HighlightBracketPair(bracketRange, matchedRange)
                GoTo RestoreCursor
            End If
        End If
    End If
    
    ' 커서 앞 문자 확인
    If hasBracketBefore Then
        Set bracketRange = rangeBefore.Duplicate
        Set matchedRange = FindMatchingBracket(bracketRange, charBefore)
        
        If Not matchedRange Is Nothing Then
            ' 괄호 쌍 하이라이트
            Call HighlightBracketPair(bracketRange, matchedRange)
            ' 하이라이트를 찾았으므로 종료
            GoTo RestoreCursor
        End If
    End If
    
    ' 커서 뒤 문자 확인 (커서 앞에 괄호가 없을 경우)
    If hasBracketAfter Then
        Set bracketRange = rangeAfter.Duplicate
        Set matchedRange = FindMatchingBracket(bracketRange, charAfter)
        
        If Not matchedRange Is Nothing Then
            ' 괄호 쌍 하이라이트
            Call HighlightBracketPair(bracketRange, matchedRange)
        End If
    End If
    
RestoreCursor:
    
    ' 원래 커서 위치로 복원
    Selection.SetRange originalRange.Start, originalRange.End
    
    isProcessingBracketMatch = False
    Exit Sub
    
ErrorHandler:
    Debug.Print "괄호 매칭 중 오류: " & Err.Description
    isProcessingBracketMatch = False
    ' 원래 커서 위치로 복원 시도
    On Error Resume Next
    If Not originalRange Is Nothing Then
        Selection.SetRange originalRange.Start, originalRange.End
    End If
End Sub

' 문자가 괄호인지 확인하는 함수
Private Function IsBracket(char As String) As Boolean
    IsBracket = (char = "(" Or char = ")" Or _
                 char = "[" Or char = "]" Or _
                 char = "{" Or char = "}")
End Function

' 여는 괄호인지 확인하는 함수
Private Function IsOpenBracket(char As String) As Boolean
    IsOpenBracket = (char = "(" Or char = "[" Or char = "{")
End Function

' 닫는 괄호인지 확인하는 함수
Private Function IsCloseBracket(char As String) As Boolean
    IsCloseBracket = (char = ")" Or char = "]" Or char = "}")
End Function

' 괄호 쌍을 찾는 함수
Private Function FindMatchingBracket(bracketRange As Range, bracketChar As String) As Range
    Dim docRange As Range
    Dim char As String
    Dim stack As Long
    Dim openBracket As String
    Dim closeBracket As String
    Dim direction As Long ' 1: 앞으로, -1: 뒤로
    Dim currentPos As Long
    
    ' 괄호 쌍 정의
    If bracketChar = "(" Then
        openBracket = "("
        closeBracket = ")"
        direction = 1 ' 여는 괄호면 앞으로
    ElseIf bracketChar = ")" Then
        openBracket = "("
        closeBracket = ")"
        direction = -1 ' 닫는 괄호면 뒤로
    ElseIf bracketChar = "[" Then
        openBracket = "["
        closeBracket = "]"
        direction = 1
    ElseIf bracketChar = "]" Then
        openBracket = "["
        closeBracket = "]"
        direction = -1
    ElseIf bracketChar = "{" Then
        openBracket = "{"
        closeBracket = "}"
        direction = 1
    ElseIf bracketChar = "}" Then
        openBracket = "{"
        closeBracket = "}"
        direction = -1
    Else
        Set FindMatchingBracket = Nothing
        Exit Function
    End If
    
    ' 스택 초기화
    stack = 1 ' 현재 괄호를 포함
    currentPos = bracketRange.Start
    
    ' 화면 업데이트 일시 중지
    Application.ScreenUpdating = False
    
    ' 괄호 검색
    Do
        ' 위치 이동
        currentPos = currentPos + direction
        
        ' 문서 범위를 벗어나면 종료
        If currentPos < 0 Or currentPos >= ActiveDocument.Content.End Then
            Exit Do
        End If
        
        ' 현재 위치의 문자 확인
        Set docRange = ActiveDocument.Range(currentPos, currentPos + 1)
        char = docRange.Text
        
        ' 같은 종류의 괄호인지 확인
        If char = openBracket Then
            stack = stack + direction
        ElseIf char = closeBracket Then
            stack = stack - direction
        End If
        
        ' 스택이 0이 되면 짝을 찾은 것
        If stack = 0 Then
            Set FindMatchingBracket = docRange.Duplicate
            Application.ScreenUpdating = True
            Exit Function
        End If
    Loop
    
    ' 짝을 찾지 못함
    Application.ScreenUpdating = True
    Set FindMatchingBracket = Nothing
End Function

' 괄호 쌍 하이라이트 함수
Private Sub HighlightBracketPair(bracket1Range As Range, bracket2Range As Range)
    On Error GoTo ErrorHandler
    
    Dim undoRecordName As String
    Dim highlightColor As Long
    
    highlightColor = RGB(173, 216, 230) ' LightBlue
    
    ' 화면 업데이트 일시 중지
    Application.ScreenUpdating = False
    
    ' 이전 하이라이트가 있으면 같은 CustomRecord 내에서 제거
    If Not previousBracketRanges Is Nothing And previousBracketRanges.Count > 0 Then
        ' 이전 하이라이트 제거 (CustomRecord 유지)
        Call RemoveBracketHighlight
    End If
    
    ' CustomRecord가 활성화되어 있지 않으면 시작
    If Not isUndoRecordActive Then
        undoRecordName = "BracketHighlight_" & Timer
        On Error Resume Next
        Application.UndoRecord.StartCustomRecord undoRecordName
        On Error GoTo ErrorHandler
        isUndoRecordActive = True
    End If
    
    ' 첫 번째 괄호의 원래 배경색 저장
    Dim originalColor1 As Long
    originalColor1 = bracket1Range.Shading.BackgroundPatternColor
    
    ' 두 번째 괄호의 원래 배경색 저장
    Dim originalColor2 As Long
    originalColor2 = bracket2Range.Shading.BackgroundPatternColor
    
    ' 첫 번째 괄호 하이라이트 (파란색)
    bracket1Range.Shading.BackgroundPatternColor = highlightColor
    
    ' 두 번째 괄호 하이라이트 (파란색)
    bracket2Range.Shading.BackgroundPatternColor = highlightColor
    
    ' 하이라이트된 범위와 원래 배경색 저장
    previousBracketRanges.Add bracket1Range.Duplicate
    previousBracketColors.Add originalColor1
    previousBracketRanges.Add bracket2Range.Duplicate
    previousBracketColors.Add originalColor2
    
    ' 화면 업데이트 재개
    Application.ScreenUpdating = True
    
    ' Undo 기록은 하이라이트가 제거될 때까지 유지 (여기서 종료하지 않음)
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "괄호 하이라이트 중 오류: " & Err.Description
    Application.ScreenUpdating = True
    On Error Resume Next
    If isUndoRecordActive Then
        Application.UndoRecord.EndCustomRecord
        isUndoRecordActive = False
    End If
End Sub

' 이전 하이라이트 제거 함수
' 항상 CustomRecord를 유지 (종료하지 않음)
Public Sub RemoveBracketHighlight()
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim highlightRange As Range
    
    ' 이전 하이라이트가 없으면 종료 (CustomRecord는 유지)
    If previousBracketRanges Is Nothing Then
        Exit Sub
    End If
    
    If previousBracketRanges.Count = 0 Then
        Exit Sub
    End If
    
    ' 화면 업데이트 일시 중지
    Application.ScreenUpdating = False
    
    ' 모든 하이라이트를 원래 배경색으로 복원 (기존 UndoRecord 내에서 실행)
    For i = 1 To previousBracketRanges.Count
        Set highlightRange = previousBracketRanges(i)
        Dim originalColor As Long
        originalColor = previousBracketColors(i)
        highlightRange.Shading.BackgroundPatternColor = originalColor
    Next i
    
    ' 컬렉션 초기화
    Set previousBracketRanges = New Collection
    Set previousBracketColors = New Collection
    
    ' 화면 업데이트 재개
    Application.ScreenUpdating = True
    
    ' CustomRecord는 항상 유지 (종료하지 않음)
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "하이라이트 제거 중 오류: " & Err.Description
    Application.ScreenUpdating = True
    ' 컬렉션 초기화
    Set previousBracketRanges = New Collection
    Set previousBracketColors = New Collection
End Sub
