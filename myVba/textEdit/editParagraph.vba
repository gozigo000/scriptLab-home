Option Explicit

' 현재 커서가 있는 "문단"을 복사합니다.
' - 선택 영역이 있으면: Word 기본 Copy처럼 선택 영역을 복사합니다.
' - 선택 영역이 없으면: 현재 커서가 포함된 문단 전체를 복사합니다.
Public Sub CopyCurrentParagraphOrSelection()
    On Error GoTo SafeExit

    Dim sel As Word.Selection
    Set sel = Application.Selection
    If sel Is Nothing Then Exit Sub

    If sel.Range.Start <> sel.Range.End Then
        sel.Copy
        Exit Sub
    End If

    Dim paraRng As Word.Range
    Set paraRng = GetCurrentParagraphRange(sel.Range)
    If paraRng Is Nothing Then Exit Sub

    paraRng.Copy

SafeExit:
End Sub

' 현재 커서가 있는 "문단"을 잘라냅니다.
' - 선택 영역이 있으면: Word 기본 Cut처럼 선택 영역을 잘라냅니다.
' - 선택 영역이 없으면: 현재 커서가 포함된 문단 전체를 잘라냅니다.
Public Sub CutCurrentParagraphOrSelection()
    On Error GoTo SafeExit

    Dim sel As Word.Selection
    Set sel = Application.Selection
    If sel Is Nothing Then Exit Sub

    If sel.Range.Start <> sel.Range.End Then
        sel.Cut
        Exit Sub
    End If

    Dim paraRng As Word.Range
    Set paraRng = GetCurrentParagraphRange(sel.Range)
    If paraRng Is Nothing Then Exit Sub

    paraRng.Cut

SafeExit:
End Sub

Private Function GetCurrentParagraphRange(ByVal rng As Word.Range) As Word.Range
    On Error GoTo ReturnNothing

    Dim paraRng As Word.Range
    Set paraRng = rng.Paragraphs(1).Range.Duplicate

    If paraRng.Start < paraRng.End Then
        ' 테이블 셀의 마지막 글자(Chr(7))는 셀 종료 마커라 잘라내면 안됨
        If paraRng.Characters.Last.Text = Chr$(7) Then
            paraRng.End = paraRng.End - 1
        End If
    End If

    Set GetCurrentParagraphRange = paraRng
    Exit Function

ReturnNothing:
    Set GetCurrentParagraphRange = Nothing
End Function
