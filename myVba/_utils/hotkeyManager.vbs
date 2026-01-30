Option Explicit

' 단축키 등록/관리 모듈
' ----------------------
' - ThisDocument 같은 곳에서 호출하는 전역 단축키 등록 헬퍼를 제공합니다.
'
' 사용 예:
'   Call RegisterHotkey("ToggleCurrWordHighlighter", wdKeyAlt, wdKeyW)
Public Sub RegisterHotkey( _
    ByVal Command As String, _
    ByVal Modifier As Long, _
    ByVal Key As Long _
)
    On Error GoTo SafeExit

    Dim keyCode As Long
    keyCode = BuildKeyCode(Modifier, Key)

    ' 매크로가 들어있는 템플릿에 키 바인딩을 저장 (문서별로 관리하기 쉬움)
    Application.CustomizationContext = ActiveDocument.AttachedTemplate

    Dim kb As Word.KeyBinding
    Set kb = Application.FindKey(keyCode)
    If Not kb Is Nothing Then
        kb.Clear
    End If

    Application.KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                                Command:=Command, _
                                KeyCode:=keyCode

SafeExit:
End Sub