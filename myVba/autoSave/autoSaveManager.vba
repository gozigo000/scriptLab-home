Option Explicit

' ======================
' autoSaveManager
' ======================
' 문서가 "Dirty(Saved=False)" 상태가 된 뒤,
' 입력이 멈춘 상태로 일정 시간(기본 5초) 지나면 자동으로 저장합니다(디바운스).
'
' 트리거:
' - clsAppEvents.appWord_WindowSelectionChange에서 
'   MaybeScheduleAutoSave(doc) 호출
'
' 제약:
' - 저장 경로가 없는 새 문서(doc.Path="")는 Save가 SaveAs UI를 
'   띄울 수 있어 스킵합니다.
'
' Public API:
' - MaybeScheduleAutoSave(doc, debounceMs)
' - CancelAutoSave([doc])

Private Const DEFAULT_DEBOUNCE_MS As Long = 5000 ' 5초
Private Const CURSOR_IDLE_REQUIRED_SEC As Double = 0.5 ' 0.5초
Private Const RETRY_DELAY_MS As Long = 500 ' 0.5초
Private Const GCS_COMPSTR As Long = &H8

#If VBA7 Then
Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr _
) As LongPtr

Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr _
) As Long

Private Declare PtrSafe Function FindWindowW Lib "user32" ( _
    ByVal lpClassName As LongPtr, _
    ByVal lpWindowName As LongPtr _
) As LongPtr

Private Declare PtrSafe Function ImmGetContext Lib "imm32" ( _
    ByVal hwnd As LongPtr _
) As LongPtr

Private Declare PtrSafe Function ImmReleaseContext Lib "imm32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hIMC As LongPtr _
) As Long

Private Declare PtrSafe Function ImmGetCompositionStringW Lib "imm32" ( _
    ByVal hIMC As LongPtr, _
    ByVal dwIndex As Long, _
    ByVal lpBuf As LongPtr, _
    ByVal dwBufLen As Long _
) As Long
#Else
Private Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long _
) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long _
) As Long

Private Declare Function FindWindowW Lib "user32" ( _
    ByVal lpClassName As Long, _
    ByVal lpWindowName As Long _
) As Long

Private Declare Function ImmGetContext Lib "imm32" ( _
    ByVal hwnd As Long _
) As Long

Private Declare Function ImmReleaseContext Lib "imm32" ( _
    ByVal hwnd As Long, _
    ByVal hIMC As Long _
) As Long

Private Declare Function ImmGetCompositionStringW Lib "imm32" ( _
    ByVal hIMC As Long, _
    ByVal dwIndex As Long, _
    ByVal lpBuf As Long, _
    ByVal dwBufLen As Long _
) As Long
#End If

#If VBA7 Then
Private gTimerId As LongPtr
#Else
Private gTimerId As Long
#End If

Private gTargetDoc As Word.Document
Private gIsAutoSaving As Boolean
Private gLastCursorDocName As String
Private gLastCursorTick As Double

Public Sub NotifyCursorActivity(ByVal Sel As Word.Selection)
    On Error GoTo SafeExit

    Dim docName As String
    On Error Resume Next
    docName = Sel.Range.Document.FullName
    If docName = "" Then docName = Sel.Range.Document.Name
    On Error GoTo SafeExit

    gLastCursorDocName = docName
    gLastCursorTick = Timer

SafeExit:
End Sub

Public Sub MaybeScheduleAutoSave( _
    ByVal doc As Word.Document, _
    Optional ByVal debounceMs As Long = DEFAULT_DEBOUNCE_MS _
)
    On Error GoTo SafeExit

    If doc Is Nothing Then Exit Sub

    ' IME(한글 조합) 중에는 자동저장 예약 자체를 하지 않음
    If IsImeComposing() Then
        Call CancelAutoSave(doc)
        Exit Sub
    End If

    ' Dirty가 아니면 예약 취소(불필요한 Save 방지)
    If doc.Saved Then
        Call CancelAutoSave(doc)
        Exit Sub
    End If

    ' SaveAs UI가 뜨는 케이스는 스킵
    If doc.ReadOnly Then Exit Sub
    If Len(doc.Path) = 0 Then Exit Sub

    If debounceMs <= 0 Then debounceMs = DEFAULT_DEBOUNCE_MS

    Set gTargetDoc = doc

    Call CancelTimer()
    gTimerId = SetTimer(0, 0, debounceMs, AddressOf AutoSaveTimerProc)

SafeExit:
End Sub

Public Sub CancelAutoSave(Optional ByVal doc As Word.Document)
    On Error GoTo SafeExit

    If Not doc Is Nothing Then
        If Not gTargetDoc Is Nothing Then
            If gTargetDoc Is doc Then
                Set gTargetDoc = Nothing
            End If
        End If
    Else
        Set gTargetDoc = Nothing
    End If

    Call CancelTimer()

SafeExit:
End Sub

#If VBA7 Then
Public Sub AutoSaveTimerProc( _
    ByVal hwnd As LongPtr, _
    ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, _
    ByVal dwTime As Long _
)
#Else
Public Sub AutoSaveTimerProc( _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal idEvent As Long, _
    ByVal dwTime As Long _
)
#End If
    ' Timer callback: 반드시 짧게, 여기서 바로 저장 실행
    Call DoAutoSaveNow()
End Sub

Private Sub DoAutoSaveNow()
    On Error GoTo SafeExit

    If gIsAutoSaving Then Exit Sub
    gIsAutoSaving = True

    Call CancelTimer()

    Dim doc As Word.Document
    Set doc = gTargetDoc
    Set gTargetDoc = Nothing

    If doc Is Nothing Then GoTo SafeExit

    If doc.Saved Then GoTo SafeExit
    If doc.ReadOnly Then GoTo SafeExit
    If Len(doc.Path) = 0 Then GoTo SafeExit

    If Not WasCursorIdleEnough(doc) Then
        Call MaybeScheduleAutoSave(doc, RETRY_DELAY_MS)
        GoTo SafeExit
    End If

    If IsImeComposing() Then
        ' IME(한글 조합) 중에는 재시도도 하지 않음
        GoTo SafeExit
    End If

    On Error Resume Next
    doc.Save
    On Error GoTo SafeExit

SafeExit:
    gIsAutoSaving = False
End Sub

Private Function WasCursorIdleEnough(ByVal doc As Word.Document) As Boolean
    On Error GoTo SafeExit

    If gLastCursorTick = 0 Then
        WasCursorIdleEnough = True
        Exit Function
    End If

    Dim docName As String
    On Error Resume Next
    docName = doc.FullName
    If docName = "" Then docName = doc.Name
    On Error GoTo SafeExit

    If docName <> "" Then
        If gLastCursorDocName <> docName Then
            WasCursorIdleEnough = True
            Exit Function
        End If
    End If

    Dim nowTick As Double
    nowTick = Timer
    If nowTick < gLastCursorTick Then gLastCursorTick = 0
    WasCursorIdleEnough = ((nowTick - gLastCursorTick) >= CURSOR_IDLE_REQUIRED_SEC)

SafeExit:
End Function

Private Function IsImeComposing() As Boolean
    On Error GoTo SafeExit

#If VBA7 Then
    Dim hwnd As LongPtr
    hwnd = GetWordHwnd()
    If hwnd = 0 Then GoTo SafeExit

    Dim hImc As LongPtr
    hImc = ImmGetContext(hwnd)
    If hImc = 0 Then GoTo SafeExit

    Dim cb As Long
    cb = ImmGetCompositionStringW(hImc, GCS_COMPSTR, 0, 0)
    Call ImmReleaseContext(hwnd, hImc)

    If cb > 0 Then IsImeComposing = True
#Else
    Dim hwnd32 As Long
    hwnd32 = GetWordHwnd()
    If hwnd32 = 0 Then GoTo SafeExit

    Dim hImc32 As Long
    hImc32 = ImmGetContext(hwnd32)
    If hImc32 = 0 Then GoTo SafeExit

    Dim cb32 As Long
    cb32 = ImmGetCompositionStringW(hImc32, GCS_COMPSTR, 0, 0)
    Call ImmReleaseContext(hwnd32, hImc32)

    If cb32 > 0 Then IsImeComposing = True
#End If

SafeExit:
End Function

#If VBA7 Then
Private Function GetWordHwnd() As LongPtr
#Else
Private Function GetWordHwnd() As Long
#End If
    On Error GoTo SafeExit

    On Error Resume Next
    GetWordHwnd = Application.ActiveWindow.Hwnd
    On Error GoTo SafeExit

    If GetWordHwnd = 0 Then
        GetWordHwnd = FindWordMainHwnd()
    End If

SafeExit:
End Function

#If VBA7 Then
Private Function FindWordMainHwnd() As LongPtr
#Else
Private Function FindWordMainHwnd() As Long
#End If
    On Error GoTo SafeExit

    ' Word 메인 윈도우 클래스명: "OpusApp"
    FindWordMainHwnd = FindWindowW(StrPtr("OpusApp"), 0)

SafeExit:
End Function

Private Sub CancelTimer()
    On Error GoTo SafeExit

    If gTimerId <> 0 Then
        Call KillTimer(0, gTimerId)
        gTimerId = 0
    End If

SafeExit:
End Sub

