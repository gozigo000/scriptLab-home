' 클래스 모듈: clsAppEvents
' ----------------------
' Application 이벤트를 처리하기 위한 클래스 모듈입니다.
' 
' 이 파일을 VBA 프로젝트에 추가하려면:
' 1. VBA 편집기에서 "삽입" > "클래스 모듈" 선택
' 2. 속성 창에서 Name을 "clsAppEvents"로 변경
' 3. 이 파일의 내용을 복사하여 붙여넣기

Public WithEvents appWord As Word.Application

' 선택 영역이 변경될 때 발생하는 이벤트
Private Sub appWord_WindowSelectionChange(ByVal Sel As Selection)

    ' bracketMatcher 모듈의 함수 호출 (괄호 매칭 및 하이라이트)
    Call HighlightBracket

    ' currWordHighlighter 모듈의 함수 호출 (현재 단어 하이라이트)
    Call HighlightCurrWord
    
    ' cursorTracker 모듈의 함수 호출 (문서별 커서 이동 히스토리 저장)
    Call WatchCursorMove(Sel)

End Sub

