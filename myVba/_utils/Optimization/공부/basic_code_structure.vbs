Option Explicit

' 🧨 백그라운드 문서 복제: 개념
' 원본 문서를 그대로 복제한 후, 복제본에서 모든 변경 작업을 수행하고, 끝나면 원본에 반영하는 패턴

' 이 패턴을 쓰면:
' UI 렌더링 부담 없음 → 화면 깜빡임 방지
' Undo 스택 부담 없음 → 대량 작업 시 속도 향상
' Range/Paragraph/Table 구조 보호 → 실수로 원본 훼손 방지
' 대용량 문서 안전 처리 가능 → 100~200페이지도 부담 최소화
' 즉, “안전 + 속도” 두 마리 토끼 잡기 전략이다.

' 기본 코드 구조
Public Sub ProcessInBackground_1()
    Dim docOriginal As Document
    Dim docCopy As Document
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As WdAlertLevel
    
    ' 원본 문서
    Set docOriginal = ActiveDocument
    
    ' 화면 갱신 OFF (이전 값 저장)
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo CleanUp

    ' 백그라운드 복제 (화면에 안 보이게 생성)
    Set docCopy = Documents.Add(Visible:=False)
    If docCopy.Windows.Count > 0 Then docCopy.Windows(1).Visible = False

    docCopy.Content.FormattedText = docOriginal.Content.FormattedText
    
    ' -------------------------
    ' 여기서 대량 작업 수행
    ' Range, Tables, Paragraphs 등 모든 변경 가능
    ' -------------------------
    ' 예시: 모든 표의 첫 번째 셀 텍스트 바꾸기
    Dim tbl As Table
    For Each tbl In docCopy.Tables
        tbl.Cell(1, 1).Range.Text = "Processed"
    Next
    
    ' 작업 끝 → 원본에 반영
    docOriginal.Content.FormattedText = docCopy.Content.FormattedText
    
    ' 백그라운드 문서 닫기
    docCopy.Close SaveChanges:=False
    
CleanUp:
    ' 화면 갱신 복구
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
End Sub