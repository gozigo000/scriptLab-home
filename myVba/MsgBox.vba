' (MARK) 자동 닫기 메시지 박스 모듈
' ----------------------
' 지정된 시간 후 자동으로 닫히는 메시지 박스를 표시하는 기능을 제공합니다.
'
' 사용 방법:
'   Call ShowAutoCloseMessageBox("메시지 내용", "제목", vbInformation, 1000)
'   timeoutMs: 밀리초 단위 (예: 1000 = 1초)

' Windows API 선언 (메시지 박스 자동 닫기용)
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal uType As Long, _
        ByVal wLanguageId As Long, _
        ByVal dwTimeout As Long) As Long
#Else
    Private Declare Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
        ByVal hwnd As Long, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal uType As Long, _
        ByVal wLanguageId As Long, _
        ByVal dwTimeout As Long) As Long
#End If

' 자동으로 닫히는 메시지 박스 표시 함수
' 매개변수:
'   message: 표시할 메시지 텍스트
'   title: 메시지 박스 제목
'   msgType: 메시지 박스 타입 (vbInformation, vbExclamation, vbCritical, vbQuestion 등)
'   timeoutMs: 자동으로 닫히기까지의 시간 (밀리초 단위)
Public Sub showMsg(message As String, title As String, msgType As Long, timeoutMs As Long)
    #If VBA7 Then
        Call MessageBoxTimeout(0, message, title, msgType, 0, timeoutMs)
    #Else
        Call MessageBoxTimeout(0, message, title, msgType, 0, timeoutMs)
    #End If
End Sub

