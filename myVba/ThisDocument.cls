' ThisDocument 모듈에 추가할 코드
' ----------------------
' 이 코드를 ThisDocument 모듈에 복사하여 붙여넣으세요.
'
' VBA 편집기에서:
' 1. 프로젝트 탐색기에서 "ThisDocument" 더블 클릭
' 2. 아래 코드를 붙여넣기

Private Sub Document_Open()
    ' currWordHighlighter 초기화
    Call InitializeCurrWordHighlighter
    ' currWordHighlighter 토글 단축키 등록
    Call RegisterHotkey("ToggleCurrWordHighlighter", wdKeyAlt, wdKeyW)
    
    ' bracketMatcher 초기화
    Call InitializeBracketMatcher
    ' bracketMatcher 토글 단축키 등록
    Call RegisterHotkey("ToggleBracketMatcher", wdKeyAlt, wdKeyB)
    
    ' cursorMoveTracker 초기화
    Call InitializeCursorMoveTracking
    ' cursorMoveTracker 토글 단축키 등록
    Call RegisterHotkey("ToggleCursorMoveTracking", wdKeyAlt, wdKeyR)

    ' 현재 문단 잘라내기
    Call RegisterHotkey("CutCurrentParagraphOrSelection", wdKeyControl, wdKeyX)

    ' 현재 문단 복사하기
    Call RegisterHotkey("CopyCurrentParagraphOrSelection", wdKeyControl, wdKeyC)
    
    ' 클래스 인스턴스 생성 및 Application 객체 연결
    Call EnsureAppEvents()
End Sub
