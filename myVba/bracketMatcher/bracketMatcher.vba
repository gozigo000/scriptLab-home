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
Public maxBracketDepth As Long ' 최대 표시 깊이 (0 = 메인 괄호만, 1 = 1단계 중첩까지, ...)

' 초기화 프로시저
Public Sub InitializeBracketMatcher()
    Set previousBracketRanges = New Collection
    Set previousBracketColors = New Collection
    isProcessingBracketMatch = False
    isUndoRecordActive = False
    maxBracketDepth = 1 ' 기본값: 1단계 중첩까지 표시
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

' 괄호의 종류를 반환하는 함수 ("()", "[]", "{}")
Private Function GetBracketType(char As String) As String
    If char = "(" Or char = ")" Then
        GetBracketType = "()"
    ElseIf char = "[" Or char = "]" Then
        GetBracketType = "[]"
    ElseIf char = "{" Or char = "}" Then
        GetBracketType = "{}"
    Else
        GetBracketType = ""
    End If
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
    Dim highlightColors() As Long
    Dim colorIndex As Long
    Dim openRange As Range
    Dim closeRange As Range
    
    ' 두 Range의 위치를 비교하여 여는 괄호와 닫는 괄호를 구분
    If bracket1Range.Start < bracket2Range.Start Then
        ' bracket1Range가 앞에 있으면 여는 괄호
        Set openRange = bracket1Range.Duplicate
        Set closeRange = bracket2Range.Duplicate
    Else
        ' bracket2Range가 앞에 있으면 여는 괄호
        Set openRange = bracket2Range.Duplicate
        Set closeRange = bracket1Range.Duplicate
    End If
    
    ' 눈에 잘 띄는 색상 팔레트 정의
    ReDim highlightColors(0 To 7)
    highlightColors(0) = RGB(255, 100, 100) ' 형광 빨강
    highlightColors(1) = RGB(100, 255, 255) ' 형광 청록
    highlightColors(2) = RGB(255, 100, 255) ' 형광 핑크
    highlightColors(3) = RGB(100, 150, 255) ' 형광 파랑
    highlightColors(4) = RGB(255, 255, 100) ' 형광 노랑
    highlightColors(5) = RGB(100, 255, 100) ' 형광 초록
    highlightColors(6) = RGB(255, 180, 50) ' 형광 주황
    highlightColors(7) = RGB(200, 100, 255) ' 형광 보라
    
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
    
    ' 메인 괄호쌍과 중첩된 모든 괄호쌍 하이라이트
    Call HighlightNestedBrackets(openRange, closeRange, highlightColors)
    
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

' 중첩된 모든 괄호쌍을 찾아 하이라이트하는 함수
' 단순한 스택 알고리즘: 여는 괄호는 스택에 추가, 닫는 괄호는 스택에서 꺼내기
' 같은 깊이의 괄호는 같은 색으로 칠함 (스택 크기를 기준으로 깊이 결정)
' 가장 바깥 괄호와 같은 종류의 괄호만 표시
Private Sub HighlightNestedBrackets(openRange As Range, closeRange As Range, colors() As Long)
    On Error GoTo ErrorHandler
    
    Dim startPos As Long
    Dim endPos As Long
    Dim currentPos As Long
    Dim char As String
    Dim docRange As Range
    Dim bracketStack As Collection ' 스택: 여는 괄호의 Range 저장
    Dim depth As Long ' 현재 깊이 (스택 크기 + 1)
    Dim mainBracketType As String ' 메인 괄호쌍의 종류
    
    startPos = openRange.Start + 1 ' 여는 괄호 다음 위치부터
    endPos = closeRange.Start ' 닫는 괄호 이전 위치까지
    
    ' 메인 괄호쌍의 종류 저장
    mainBracketType = GetBracketType(openRange.Text)
    
    ' 메인 괄호쌍 하이라이트 (깊이 0)
    Dim originalColor1 As Long
    Dim originalColor2 As Long
    originalColor1 = openRange.Shading.BackgroundPatternColor
    originalColor2 = closeRange.Shading.BackgroundPatternColor
    
    openRange.Shading.BackgroundPatternColor = colors(0 Mod (UBound(colors) + 1))
    closeRange.Shading.BackgroundPatternColor = colors(0 Mod (UBound(colors) + 1))
    
    previousBracketRanges.Add openRange.Duplicate
    previousBracketColors.Add originalColor1
    previousBracketRanges.Add closeRange.Duplicate
    previousBracketColors.Add originalColor2
    
    ' 스택 초기화
    Set bracketStack = New Collection
    
    ' 괄호 범위 내에서 모든 괄호쌍 찾기
    currentPos = startPos
    Do While currentPos < endPos
        Set docRange = ActiveDocument.Range(currentPos, currentPos + 1)
        char = docRange.Text
        
        ' 괄호인지 확인하고, 메인 괄호쌍과 같은 종류인지 확인
        If IsBracket(char) And GetBracketType(char) = mainBracketType Then
            If IsOpenBracket(char) Then
                ' 여는 괄호: 스택에 추가 (메인 괄호쌍과 같은 종류만)
                bracketStack.Add docRange.Duplicate
            ElseIf IsCloseBracket(char) Then
                ' 닫는 괄호: 스택에서 꺼내기
                If bracketStack.Count > 0 Then
                    ' 스택의 맨 위(마지막) 요소 가져오기
                    Dim openBracketRange As Range
                    Dim closeBracketRange As Range
                    Dim stackChar As String
                    
                    Set openBracketRange = bracketStack(bracketStack.Count)
                    stackChar = openBracketRange.Text
                    
                    ' 같은 종류의 괄호인지 확인 (이미 mainBracketType과 같은지만 확인했지만, 스택의 괄호와도 확인)
                    If GetBracketType(stackChar) = GetBracketType(char) Then
                        ' 짝을 찾음: 깊이 확인 후 하이라이트 적용
                        ' 현재 깊이는 스택 크기 (제거하기 전, 여는 괄호가 추가되기 전의 깊이 + 1)
                        ' 메인 괄호쌍이 깊이 0이므로, 첫 번째 중첩은 깊이 1
                        depth = bracketStack.Count
                        
                        ' 최대 깊이를 초과하지 않으면 하이라이트 적용
                        If depth <= maxBracketDepth Then
                            Set closeBracketRange = docRange.Duplicate
                            
                            ' 원래 배경색 저장
                            Dim origColorOpen As Long
                            Dim origColorClose As Long
                            origColorOpen = openBracketRange.Shading.BackgroundPatternColor
                            origColorClose = closeBracketRange.Shading.BackgroundPatternColor
                            
                            ' 하이라이트 적용 (깊이를 색상 인덱스로 사용)
                            openBracketRange.Shading.BackgroundPatternColor = colors(depth Mod (UBound(colors) + 1))
                            closeBracketRange.Shading.BackgroundPatternColor = colors(depth Mod (UBound(colors) + 1))
                            
                            ' 저장
                            previousBracketRanges.Add openBracketRange.Duplicate
                            previousBracketColors.Add origColorOpen
                            previousBracketRanges.Add closeBracketRange.Duplicate
                            previousBracketColors.Add origColorClose
                        End If
                        
                        ' 스택에서 제거 (깊이와 관계없이 제거해야 다음 괄호 매칭이 올바르게 됨)
                        bracketStack.Remove bracketStack.Count
                    End If
                End If
            End If
        End If
        
        currentPos = currentPos + 1
    Loop
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "중첩 괄호 하이라이트 중 오류: " & Err.Description
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
