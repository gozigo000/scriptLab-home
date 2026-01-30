Public gAppEvents As clsAppEvents

' Word 시작/템플릿 로드 시 자동 실행
' AutoExec는 Word가 자동으로 호출하는 "특수 이름 매크로"
' - 호출 주체: Word 자체
' - 호출 시점: 보통 전역 템플릿(Normal.dotm)이나 추가기능(.dotm, Add-in)이 
'   로드될 때 (문서가 열릴 때마다가 아니라 "Word/템플릿 로드 타이밍" 쪽에 가까운 시점)
' - 주의: 문서(.docm)에만 들어있는 경우에는 기대한 타이밍에 안 불릴 수 있어서, 
'   지금처럼 ThisDocument.Document_Open에서 EnsureAppEvents()를 같이 
'   호출해두면 안정적으로 이벤트 연결이 됩니다.
Public Sub AutoExec()
    Call EnsureAppEvents()
End Sub

' Application 이벤트 핸들러를 1회만 생성/연결합니다.
' 여러 문서가 열려도 중복 핸들러가 생기지 않게 전역으로 유지합니다.
Public Sub EnsureAppEvents()
    On Error GoTo SafeExit

    If gAppEvents Is Nothing Then
        Set gAppEvents = New clsAppEvents
    End If

    Set gAppEvents.appWord = Word.Application

SafeExit:
End Sub

