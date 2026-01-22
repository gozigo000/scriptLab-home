Option Explicit

' ============================================================
' 모듈: undoRecordUtils
' 역할: Word UndoRecord(CustomRecord) 유틸
' Public API:
'   - BeginCustomUndoRecord(recordName = "CustomUndoRecord"): CustomRecord 시작
'   - EndCustomUndoRecord(): CustomRecord 종료
' ============================================================
'
' 사용 예:
'   Call BeginCustomUndoRecord() ' 또는 Call BeginCustomUndoRecord("MyAction")
'   ' ... document changes ...
'   Call EndCustomUndoRecord()

Private mUndoRecord As Object
Private mDepth As Long

' Custom Undo Record를 시작합니다(가능한 경우).
' - 내부적으로 중첩 호출을 지원합니다(Depth 기반).
' - Word 버전/환경에 따라 Application.UndoRecord가 없을 수 있으므로
'   실패 시 조용히 무시합니다(기존 동작 유지).
Public Sub BeginCustomUndoRecord(Optional ByVal recordName As String = "CustomUndoRecord")
    If recordName = "" Then recordName = "CustomUndoRecord"
    
    mDepth = mDepth + 1
    If mDepth > 1 Then Exit Sub
    
    On Error Resume Next
    Set mUndoRecord = Application.UndoRecord
    If Err.Number = 0 Then mUndoRecord.StartCustomRecord recordName
    If Err.Number <> 0 Then Set mUndoRecord = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

' BeginCustomUndoRecord로 시작한 UndoRecord를 종료합니다.
Public Sub EndCustomUndoRecord()
    If mDepth <= 0 Then Exit Sub
    
    mDepth = mDepth - 1
    If mDepth > 0 Then Exit Sub
    
    On Error Resume Next
    If Not mUndoRecord Is Nothing Then mUndoRecord.EndCustomRecord
    Set mUndoRecord = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

