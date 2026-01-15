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
    'Call InitializeSelectionAutoSearch
    
    ' bracketMatcher 초기화
    Call InitializeBracketMatcher
    ' bracketMatcher 토글 단축키 등록 (Alt+B)
    Call RegisterHotkey("ToggleBracketMatcher", BuildKeyCode(wdKeyAlt, wdKeyB))
    
    ' 클래스 인스턴스 생성 및 Application 객체 연결
    ' 이 클래스는 WindowSelectionChange 이벤트를 처리합니다.
    Set myAppEvents = New clsAppEvents
    Set myAppEvents.appWord = Word.Application
End Sub


' 단축키 등록 함수
' - Command: 실행할 매크로 이름(문자열)
' - KeyCode: BuildKeyCode(wdKeyAlt, wdKeyB) 같은 값
Private Sub RegisterHotkey(ByVal Command As String, ByVal KeyCode As Long)
    On Error GoTo SafeExit
    
    ' 매크로가 들어있는 템플릿에 키 바인딩을 저장 (문서별로 관리하기 쉬움)
    Application.CustomizationContext = ActiveDocument.AttachedTemplate
    
    Dim kb As KeyBinding
    Set kb = Application.FindKey(KeyCode)
    If Not kb Is Nothing Then
        kb.Clear
    End If
    
    Application.KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                                Command:=Command, _
                                KeyCode:=KeyCode
    
SafeExit:
End Sub
