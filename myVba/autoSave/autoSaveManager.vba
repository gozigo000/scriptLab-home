Option Explicit

' ======================
' autoSaveManager
' ======================
' 문서가 "Dirty(Saved=False)" 상태가 된 뒤,
' 입력이 멈춘 상태로 일정 시간(기본 1초) 지나면 자동으로 저장합니다(디바운스).
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

Private Const DEFAULT_DEBOUNCE_MS As Long = 1000

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
#End If

#If VBA7 Then
Private gTimerId As LongPtr
#Else
Private gTimerId As Long
#End If

Private gTargetDoc As Word.Document
Private gIsAutoSaving As Boolean

Public Sub MaybeScheduleAutoSave( _
    ByVal doc As Word.Document, _
    Optional ByVal debounceMs As Long = DEFAULT_DEBOUNCE_MS _
)
    On Error GoTo SafeExit

    If doc Is Nothing Then Exit Sub

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

    On Error Resume Next
    doc.Save
    On Error GoTo SafeExit

SafeExit:
    gIsAutoSaving = False
End Sub

Private Sub CancelTimer()
    On Error GoTo SafeExit

    If gTimerId <> 0 Then
        Call KillTimer(0, gTimerId)
        gTimerId = 0
    End If

SafeExit:
End Sub

