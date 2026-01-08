' ThisDocument 모듈에 추가할 코드
' ----------------------
' 이 코드를 ThisDocument 모듈에 복사하여 붙여넣으세요.
'
' VBA 편집기에서:
' 1. 프로젝트 탐색기에서 "ThisDocument" 더블 클릭
' 2. 아래 코드를 붙여넣기
'
' 이 코드는 다음 기능들을 초기화합니다:
' - selectionBold: 선택 영역 자동 볼드 처리
' - selectionAutoSearch: 선택 영역 자동 검색 및 하이라이트

Dim myAppEvents As clsAppEvents

Private Sub Document_Open()
    ' selectionAutoSearch 초기화
    Call InitializeSelectionAutoSearch
    
    ' 클래스 인스턴스 생성 및 Application 객체 연결
    ' 이 클래스는 WindowSelectionChange 이벤트를 처리하여
    ' selectionBold와 selectionAutoSearch 기능을 모두 실행합니다.
    Set myAppEvents = New clsAppEvents
    Set myAppEvents.appWord = Word.Application
End Sub

