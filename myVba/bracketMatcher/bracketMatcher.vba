' (MARK) 괄호 매칭 및 하이라이트 섹션
' ----------------------
' 커서가 괄호 옆에 있을 때 해당 괄호의 짝을 찾아 하이라이트하는 기능
'
' 사용 방법:
' 1. 이 모듈을 Word VBA 프로젝트에 추가합니다.
' 2. clsAppEvents 클래스 모듈을 수정하여 HighlightBracket 함수를 호출하도록 합니다.
'    (clsAppEvents.cls 파일의 appWord_WindowSelectionChange 이벤트에 추가)
' 3. ThisDocument 모듈에서 Document_Open 이벤트에 InitializeBracketMatcher 호출 추가
' 

' 모듈 레벨 변수
Public isBracketMatcherEnabled As Boolean ' 기능 ON/OFF 토글 (True=활성화)
Public maxBracketDepth As Long ' 최대 표시 깊이 (0 = 메인 괄호만, 1 = 1단계 중첩까지, ...)
Public isUndoRecordActive As Boolean ' UndoRecord가 활성화되어 있는지 추적
Public previousBracketRanges As Collection ' 이전에 하이라이트된 괄호 범위들 저장
Public previousBracketColors As Collection ' 이전 괄호의 원래 배경색 저장
Public previousOperatorRanges As Collection ' 이전에 빨강 처리된 연산자 범위들 저장
Public previousOperatorColors As Collection ' 이전 연산자의 원래 글자색 저장
Public isProcessingBracketMatch As Boolean ' 무한루프 방지 플래그

' 초기화 프로시저
Public Sub InitializeBracketMatcher()
    isBracketMatcherEnabled = False ' 초기 상태: 비활성화
    maxBracketDepth = 1 ' 기본값: 1단계 중첩까지 표시
    isUndoRecordActive = False
    Set previousBracketRanges = New Collection
    Set previousBracketColors = New Collection
    Set previousOperatorRanges = New Collection
    Set previousOperatorColors = New Collection
    isProcessingBracketMatch = False
End Sub

' 기능 토글 매크로 (수동 실행용)
' - 실행할 때마다 ON/OFF가 바뀜
Public Sub ToggleBracketMatcher()
    Call EnsureBracketMatcherInitialized
    
    isBracketMatcherEnabled = Not isBracketMatcherEnabled
    
    If isBracketMatcherEnabled Then
        ' 켤 때 현재 커서 위치에서 괄호 매칭 및 하이라이트
        On Error Resume Next
        Call HighlightBracket
        On Error GoTo 0
    Else
        ' 끌 때는 즉시 하이라이트 정리 + UndoRecord 종료
        On Error Resume Next
        Call RemoveBracketHighlight
        If isUndoRecordActive Then
            Application.UndoRecord.EndCustomRecord
            isUndoRecordActive = False
        End If
        On Error GoTo 0
    End If
    
    Debug.Print "BracketMatcher Enabled = " & CStr(isBracketMatcherEnabled)
End Sub

' 필요 시 초기화 (컬렉션이 Nothing이면 초기화)
Private Sub EnsureBracketMatcherInitialized()
    If previousBracketRanges Is Nothing Or previousBracketColors Is Nothing Or _
       previousOperatorRanges Is Nothing Or previousOperatorColors Is Nothing Then
        Call InitializeBracketMatcher
    End If
End Sub

' 괄호 매칭 및 하이라이트 함수
' 이 함수는 clsAppEvents 클래스 모듈에서 appWord_WindowSelectionChange 이벤트로 호출됨
Public Sub HighlightBracket()
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

    ' 기능이 꺼져있으면 하이라이트만 정리하고 종료
    If Not isBracketMatcherEnabled Then
        On Error Resume Next
        Call RemoveBracketHighlight
        If isUndoRecordActive Then
            Application.UndoRecord.EndCustomRecord
            isUndoRecordActive = False
        End If
        On Error GoTo 0
        Exit Sub
    End If
    
    ' 처리 중 플래그 설정
    isProcessingBracketMatch = True
    
    On Error GoTo ErrorHandler
    
    Dim cursorPos As Long
    Dim originalRange As Range
    Dim searchBounds As Range
    Dim openEnclosing As Range
    Dim closeEnclosing As Range
    
    ' 현재 커서 위치 저장
    Set originalRange = Selection.Range.Duplicate
    cursorPos = Selection.Start
    
    ' 괄호 검색 범위 제한: 현재 문단 기준 위/아래 1개 문단까지만
    Set searchBounds = GetBracketSearchBoundsAroundCursor(1)
    
    ' 이전 하이라이트 제거
    Call RemoveBracketHighlight

    ' 요구사항: 커서가 괄호 "옆"에 있더라도, 괄호쌍 "내부"가 아니면 하이라이트하지 않음.
    ' 따라서 항상 "내부인지"만 판정해서, 내부일 때만 가장 안쪽 괄호쌍을 하이라이트한다.
    If TryFindEnclosingBracketPair(cursorPos, searchBounds, openEnclosing, closeEnclosing) Then
        Call HighlightBracketPair(openEnclosing, closeEnclosing)
        GoTo RestoreCursor
    End If
    
    ' 괄호도 없고, 괄호 내부도 아니면 CustomRecord 종료
    If isUndoRecordActive Then
        On Error Resume Next
        Application.UndoRecord.EndCustomRecord
        On Error GoTo ErrorHandler
        isUndoRecordActive = False
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

' 괄호 쌍을 찾는 함수 (검색 범위를 bounds로 제한)
Private Function FindMatchingBracketInBounds(bracketRange As Range, bracketChar As String, bounds As Range) As Range
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
        Set FindMatchingBracketInBounds = Nothing
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
        
        ' bounds 범위를 벗어나면 종료
        If currentPos < bounds.Start Or currentPos >= bounds.End Then
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
            Set FindMatchingBracketInBounds = docRange.Duplicate
            Application.ScreenUpdating = True
            Exit Function
        End If
    Loop
    
    ' 짝을 찾지 못함
    Application.ScreenUpdating = True
    Set FindMatchingBracketInBounds = Nothing
End Function

' 현재 커서 문단 기준 위/아래 maxParagraphs개 문단까지의 Range를 반환
Private Function GetBracketSearchBoundsAroundCursor(ByVal maxParagraphs As Long) As Range
    Dim basePara As Paragraph
    Dim startPara As Paragraph
    Dim endPara As Paragraph
    Dim i As Long
    
    If Selection Is Nothing Or Selection.Range Is Nothing Then
        Set GetBracketSearchBoundsAroundCursor = ActiveDocument.Content
        Exit Function
    End If
    
    If Selection.Range.Paragraphs.Count = 0 Then
        Set GetBracketSearchBoundsAroundCursor = ActiveDocument.Content
        Exit Function
    End If
    
    Set basePara = Selection.Range.Paragraphs(1)
    Set startPara = basePara
    Set endPara = basePara
    
    For i = 1 To maxParagraphs
        If startPara.Previous Is Nothing Then Exit For
        Set startPara = startPara.Previous
    Next i
    
    For i = 1 To maxParagraphs
        If endPara.Next Is Nothing Then Exit For
        Set endPara = endPara.Next
    Next i
    
    Set GetBracketSearchBoundsAroundCursor = ActiveDocument.Range(startPara.Range.Start, endPara.Range.End)
End Function

' 커서가 어떤 괄호쌍 내부에 있을 때, 가장 가까운 상위(=가장 안쪽) 괄호쌍을 찾는다.
' - bounds 범위 안에서만 탐색
' - 찾으면 True + openRange/closeRange 반환
Private Function TryFindEnclosingBracketPair(ByVal cursorPos As Long, ByVal bounds As Range, ByRef openRange As Range, ByRef closeRange As Range) As Boolean
    On Error GoTo ErrorHandler
    
    Dim pos As Long
    Dim ch As String
    Dim r As Range
    
    Dim stackChars As Collection
    Dim stackPos As Collection
    Set stackChars = New Collection
    Set stackPos = New Collection
    
    ' 1) bounds.Start ~ cursorPos-1 까지 스캔하며 "아직 닫히지 않은" 가장 안쪽 여는 괄호를 찾는다.
    pos = bounds.Start
    Do While pos < cursorPos And pos < bounds.End
        Set r = ActiveDocument.Range(pos, pos + 1)
        ch = r.Text
        
        If IsOpenBracket(ch) Then
            stackChars.Add ch
            stackPos.Add pos
        ElseIf IsCloseBracket(ch) Then
            If stackChars.Count > 0 Then
                ' top과 종류가 맞으면 pop
                If GetBracketType(stackChars(stackChars.Count)) = GetBracketType(ch) Then
                    stackChars.Remove stackChars.Count
                    stackPos.Remove stackPos.Count
                End If
            End If
        End If
        
        pos = pos + 1
    Loop
    
    If stackChars.Count = 0 Then
        TryFindEnclosingBracketPair = False
        Exit Function
    End If
    
    Dim openCh As String
    Dim openPos As Long
    Dim openBracket As String
    Dim closeBracket As String
    Dim depth As Long
    
    openCh = stackChars(stackChars.Count)
    openPos = CLng(stackPos(stackPos.Count))
    
    If openCh = "(" Then
        openBracket = "("
        closeBracket = ")"
    ElseIf openCh = "[" Then
        openBracket = "["
        closeBracket = "]"
    ElseIf openCh = "{" Then
        openBracket = "{"
        closeBracket = "}"
    Else
        TryFindEnclosingBracketPair = False
        Exit Function
    End If
    
    Set openRange = ActiveDocument.Range(openPos, openPos + 1)
    
    ' 2) cursorPos ~ bounds.End-1 까지 스캔하며 해당 여는 괄호의 매칭 닫는 괄호를 찾는다.
    depth = 1
    pos = cursorPos
    Do While pos < bounds.End
        Set r = ActiveDocument.Range(pos, pos + 1)
        ch = r.Text
        
        If ch = openBracket Then
            depth = depth + 1
        ElseIf ch = closeBracket Then
            depth = depth - 1
            If depth = 0 Then
                Set closeRange = r.Duplicate
                ' cursorPos가 여는 괄호와 닫는 괄호 사이에 있으면 내부로 취급
                ' - 여는 괄호 바로 뒤(openPos+1)는 내부
                ' - 닫는 괄호 바로 앞(closePos)는 내부
                TryFindEnclosingBracketPair = True
                Exit Function
            End If
        End If
        
        pos = pos + 1
    Loop
    
    TryFindEnclosingBracketPair = False
    Exit Function
    
ErrorHandler:
    TryFindEnclosingBracketPair = False
End Function

' 괄호 쌍 하이라이트 함수
Private Sub HighlightBracketPair(bracket1Range As Range, bracket2Range As Range)
    On Error GoTo ErrorHandler
    
    Dim undoRecordName As String
    Dim highlightColors() As Long
    Dim colorIndex As Long
    Dim openRange As Range
    Dim closeRange As Range
    Dim innerRange As Range
    Dim innerText As String
    
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

    ' 괄호 내부가 "숫자"만이면 하이라이트하지 않음
    ' (예: (123), ( 2026 ), (12 34) 등 - 공백/개행은 무시)
    If openRange.End <= closeRange.Start Then
        Set innerRange = ActiveDocument.Range(openRange.End, closeRange.Start)
        innerText = innerRange.Text
        If IsDigitOnly(innerText) Then
            ' 이전 하이라이트는 이미 HighlightBracket에서 제거됨.
            ' 현재는 하이라이트를 하지 않으므로 CustomRecord가 열려있으면 종료.
            If isUndoRecordActive Then
                On Error Resume Next
                Application.UndoRecord.EndCustomRecord
                On Error GoTo ErrorHandler
                isUndoRecordActive = False
            End If
            Exit Sub
        End If
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

    ' 깊이 0(메인 괄호 내부)에서만 특정 연산자/기호를 빨강 글자색으로 표시
    Call HighlightOperatorsAtDepth0(openRange, closeRange)
    
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

' 메인(깊이 0) 괄호 내부에서만 특정 문자/연산자의 글자색을 빨강으로 표시
' - 대상: &&, ||, +, ',', ?, :, ;
' - "깊이 0" = 메인 괄호쌍 내부이되, 같은 종류 괄호로 중첩된 구간(깊이>=1) 제외
Private Sub HighlightOperatorsAtDepth0(openRange As Range, closeRange As Range)
    On Error GoTo ErrorHandler
    
    Dim startPos As Long
    Dim endPos As Long
    Dim pos As Long
    Dim depth As Long
    Dim mainType As String
    Dim r As Range
    Dim ch As String
    
    startPos = openRange.End
    endPos = closeRange.Start
    
    If startPos >= endPos Then Exit Sub
    
    mainType = GetBracketType(openRange.Text)
    depth = 0
    pos = startPos
    
    Do While pos < endPos
        Set r = ActiveDocument.Range(pos, pos + 1)
        ch = r.Text
        
        ' 같은 종류 괄호의 중첩 깊이 추적
        If IsBracket(ch) And GetBracketType(ch) = mainType Then
            If IsOpenBracket(ch) Then
                depth = depth + 1
            ElseIf IsCloseBracket(ch) Then
                If depth > 0 Then depth = depth - 1
            End If
            
            pos = pos + 1
            GoTo ContinueLoop
        End If
        
        If depth = 0 Then
            ' && 처리 (두 글자)
            If ch = "&" And pos + 1 < endPos Then
                Dim r2 As Range
                Set r2 = ActiveDocument.Range(pos + 1, pos + 2)
                If r2.Text = "&" Then
                    Call ApplyRedFontAndRemember(r)
                    Call ApplyRedFontAndRemember(r2)
                    pos = pos + 2
                    GoTo ContinueLoop
                End If
            End If
            
            ' || 처리 (두 글자, 중간 공백 허용: "| |"도 처리)
            If ch = "|" And pos + 1 < endPos Then
                Dim pos2 As Long
                Dim rPipe2 As Range
                
                pos2 = pos + 1
                Do While pos2 < endPos
                    Dim rTmp As Range
                    Set rTmp = ActiveDocument.Range(pos2, pos2 + 1)
                    If Not IsWhitespaceChar(rTmp.Text) Then Exit Do
                    pos2 = pos2 + 1
                Loop
                
                If pos2 < endPos Then
                    Set rPipe2 = ActiveDocument.Range(pos2, pos2 + 1)
                    If rPipe2.Text = "|" Then
                        Call ApplyRedFontAndRemember(r)
                        Call ApplyRedFontAndRemember(rPipe2)
                        pos = pos2 + 1
                        GoTo ContinueLoop
                    End If
                End If
            End If
            
            ' 단일 문자 처리
            If ch = "+" Or ch = "-" Or ch = "," Or ch = "?" Or ch = ":" Or ch = ";" Then
                Call ApplyRedFontAndRemember(r)
            End If
        End If
        
        pos = pos + 1
ContinueLoop:
    Loop
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "연산자 빨강 표시 중 오류: " & Err.Description
End Sub

' 공백 문자 판정 (|| 사이 공백 허용용)
Private Function IsWhitespaceChar(ByVal ch As String) As Boolean
    If ch = " " Or ch = vbTab Or ch = ChrW$(160) Then
        IsWhitespaceChar = True
    Else
        IsWhitespaceChar = False
    End If
End Function

' 글자색을 빨강으로 바꾸고, 원래 색을 복원할 수 있도록 저장
Private Sub ApplyRedFontAndRemember(targetRange As Range)
    On Error GoTo ErrorHandler
    
    If previousOperatorRanges Is Nothing Then Set previousOperatorRanges = New Collection
    If previousOperatorColors Is Nothing Then Set previousOperatorColors = New Collection
    
    Dim origColor As Long
    origColor = targetRange.Font.Color
    
    targetRange.Font.Color = RGB(255, 0, 0)
    
    previousOperatorRanges.Add targetRange.Duplicate
    previousOperatorColors.Add origColor
    
    Exit Sub
    
ErrorHandler:
    ' 실패해도 전체 하이라이트는 계속
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
    Dim opRange As Range
    
    ' 이전 하이라이트가 전혀 없으면 종료 (CustomRecord는 유지)
    Dim hasAny As Boolean
    hasAny = False
    If Not previousBracketRanges Is Nothing Then
        If previousBracketRanges.Count > 0 Then hasAny = True
    End If
    If Not previousOperatorRanges Is Nothing Then
        If previousOperatorRanges.Count > 0 Then hasAny = True
    End If
    If Not hasAny Then Exit Sub
    
    ' 화면 업데이트 일시 중지
    Application.ScreenUpdating = False
    
    ' 모든 하이라이트를 원래 배경색으로 복원 (기존 UndoRecord 내에서 실행)
    If Not previousBracketRanges Is Nothing Then
        For i = 1 To previousBracketRanges.Count
            Set highlightRange = previousBracketRanges(i)
            Dim originalColor As Long
            originalColor = previousBracketColors(i)
            highlightRange.Shading.BackgroundPatternColor = originalColor
        Next i
    End If
    
    ' 연산자 글자색 복원
    If Not previousOperatorRanges Is Nothing Then
        For i = 1 To previousOperatorRanges.Count
            Set opRange = previousOperatorRanges(i)
            Dim origFontColor As Long
            origFontColor = previousOperatorColors(i)
            opRange.Font.Color = origFontColor
        Next i
    End If
    
    ' 컬렉션 초기화
    Set previousBracketRanges = New Collection
    Set previousBracketColors = New Collection
    Set previousOperatorRanges = New Collection
    Set previousOperatorColors = New Collection
    
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
    Set previousOperatorRanges = New Collection
    Set previousOperatorColors = New Collection
End Sub

' 괄호 내부 문자열이 "숫자"로만 구성되어 있는지 확인 (공백/개행은 무시)
Private Function IsDigitOnly(ByVal s As String) As Boolean
    Dim t As String
    Dim re As Object
    
    ' 공백류 제거 (스페이스/탭/개행)
    t = s
    t = Replace(t, " ", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    
    If Len(t) = 0 Then
        IsDigitOnly = False
        Exit Function
    End If
    
    ' VBScript 정규식 (late binding)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "^[0-9]+$"
    
    IsDigitOnly = re.Test(t)
End Function