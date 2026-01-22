Option Explicit

' ============================================================
' 모듈: tableHelper
' 역할: Word 문서의 표(Table) 관련 공용 유틸
' - Range가 표 안인지 판정
' - 표를 만나면 표를 "통째로" 건너뛰어 이전/다음 문단으로 점프
' ============================================================

' Range가 표 내부인지 판정
Public Function IsRangeInTable(ByVal rng As Range) As Boolean
    On Error GoTo SafeExit
    If rng Is Nothing Then GoTo SafeExit
    IsRangeInTable = CBool(rng.Information(wdWithInTable))
    Exit Function
SafeExit:
    IsRangeInTable = False
End Function

' p가 표 안이면 표 밖(이전/다음) 문단으로 점프해서 반환
' direction: -1(위/이전 방향으로 표 탈출), +1(아래/다음 방향으로 표 탈출)
Public Function GetParagraphOutsideTable( _
    ByVal doc As Document, _
    ByVal p As Paragraph, _
    ByVal direction As Long _
) As Paragraph
    On Error GoTo SafeExit
    
    If doc Is Nothing Then GoTo SafeExit
    If p Is Nothing Then GoTo SafeExit
    
    Dim cur As Paragraph
    Set cur = p
    
    Dim guard As Long
    guard = 0
    
    Do While Not cur Is Nothing
        If Not IsRangeInTable(cur.Range) Then
            Set GetParagraphOutsideTable = cur
            Exit Function
        End If
        
        If cur.Range.Tables.Count <= 0 Then Exit Do
        
        Dim tbl As Table
        Set tbl = cur.Range.Tables(1)
        
        Dim r As Range
        Set r = tbl.Range.Duplicate
        
        If direction < 0 Then
            r.Collapse wdCollapseStart
            If r.Start <= doc.Content.Start Then Exit Do
            r.Start = r.Start - 1
            r.End = r.Start
        Else
            r.Collapse wdCollapseEnd
            If r.End >= doc.Content.End Then Exit Do
            r.Start = r.End + 1
            r.End = r.Start
        End If
        
        If r.Start < doc.Content.Start Or r.Start > doc.Content.End Then Exit Do
        Set cur = r.Paragraphs(1)
        
        guard = guard + 1
        If guard > 50 Then Exit Do
    Loop
    
SafeExit:
    Set GetParagraphOutsideTable = Nothing
End Function

Public Function GetPreviousParagraphSkippingTables( _
    ByVal doc As Document, _
    ByVal p As Paragraph _
) As Paragraph
    On Error GoTo SafeExit
    
    If doc Is Nothing Then GoTo SafeExit
    If p Is Nothing Then GoTo SafeExit
    
    Dim prevP As Paragraph
    Set prevP = p.Previous
    If prevP Is Nothing Then
        Set GetPreviousParagraphSkippingTables = Nothing
        Exit Function
    End If
    
    If IsRangeInTable(prevP.Range) Then
        Set GetPreviousParagraphSkippingTables = GetParagraphOutsideTable(doc, prevP, -1)
    Else
        Set GetPreviousParagraphSkippingTables = prevP
    End If
    Exit Function
    
SafeExit:
    Set GetPreviousParagraphSkippingTables = Nothing
End Function

Public Function GetNextParagraphSkippingTables( _
    ByVal doc As Document, _
    ByVal p As Paragraph _
) As Paragraph
    On Error GoTo ReturnNothing
    
    If doc Is Nothing Then GoTo ReturnNothing
    If p Is Nothing Then GoTo ReturnNothing
    
    Dim nextP As Paragraph
    Set nextP = p.Next
    If nextP Is Nothing Then GoTo ReturnNothing
    
    If IsRangeInTable(nextP.Range) Then
        Set GetNextParagraphSkippingTables = GetParagraphOutsideTable(doc, nextP, +1)
    Else
        Set GetNextParagraphSkippingTables = nextP
    End If
    Exit Function
    
ReturnNothing:
    Set GetNextParagraphSkippingTables = Nothing
End Function

