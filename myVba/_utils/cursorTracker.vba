Option Explicit

' (MARK) 커서 위치 정보 알림 매크로
' ----------------------
' 현재 커서(Selection) 기준으로 다음 정보를 알림창으로 표시합니다.
' - 페이지 번호 / 전체 페이지 수
' - 문단 번호(문서 기준) / 문단 텍스트(앞부분)
' - 현재 단어(커서 위치 기준)
' - 목차(TOC) 영역 여부
' - 북마크 포함 여부(해당 커서가 속한 북마크 이름들)
' - 하이퍼링크 포함 여부(주소/서브주소/표시텍스트)
'
' 사용:
' - 매크로: ShowCursorLocationInfo 실행

Public Sub ShowCursorLocationInfo()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim selRng As Range
    Dim msg As String
    
    Set doc = ActiveDocument
    Set selRng = Selection.Range.Duplicate
    
    ' ===== 페이지 =====
    Dim pageNo As Long
    Dim numPages As Long
    pageNo = Selection.Information(wdActiveEndPageNumber)
    
    ' ===== 문단 =====
    ' Dim paraIdx As Long
    Dim paraText As String
    ' paraIdx = GetParagraphIndexInDocument(doc, selRng)
    paraText = GetCurrentParagraphPreview(selRng, 120)
    
    ' ===== 영역 제목(탐색창/목차처럼 보이는 제목) =====
    Dim headingTitle As String
    headingTitle = GetNearestHeadingTitle(selRng, 140)
    
    ' ===== 단어 =====
    Dim wordText As String
    wordText = GetWordAtCursor(selRng)
    
    ' ===== 북마크 =====
    Dim bmNames As String
    bmNames = GetBookmarkNamesAtRange(doc, selRng, 15)
    
    ' ===== 하이퍼링크 =====
    Dim linkInfo As String
    linkInfo = GetHyperlinkInfoAtRange(selRng, 5)
    
    msg = ""
    msg = msg & "페이지: " & pageNo & vbCrLf
    ' msg = msg & "문단: " & paraIdx & vbCrLf
    msg = msg & "영역 제목: " & IIf(headingTitle = "", "(없음)", headingTitle) & vbCrLf
    msg = msg & "현재 단어: " & IIf(wordText = "", "(없음)", """" & wordText & """") & vbCrLf
    msg = msg & "북마크: " & IIf(bmNames = "", "(없음)", bmNames) & vbCrLf
    msg = msg & "하이퍼링크: " & IIf(linkInfo = "", "(없음)", vbCrLf & linkInfo) & vbCrLf
    msg = msg & vbCrLf & "문단 미리보기:" & vbCrLf & paraText
    
    ' 기본 MsgBox (원하면 showMsg로 교체 가능)
    VBA.MsgBox msg, vbInformation, "커서 위치 정보"
    Exit Sub
    
ErrorHandler:
    VBA.MsgBox "오류: " & Err.Description, vbCritical, "커서 위치 정보"
End Sub

' ======================
' Helpers
' ======================

' 문서 기준 문단 인덱스(1-based). 계산 실패 시 0.
Private Function GetParagraphIndexInDocument(ByVal doc As Document, ByVal rng As Range) As Long
    On Error GoTo SafeExit
    
    Dim probe As Range
    Set probe = doc.Range(0, rng.Start)
    
    ' rng.Start가 문서 첫 위치면 Paragraphs.Count가 0이 될 수 있어 보정
    If probe.End <= 0 Then
        GetParagraphIndexInDocument = 1
    Else
        GetParagraphIndexInDocument = probe.Paragraphs.Count
        If GetParagraphIndexInDocument < 1 Then GetParagraphIndexInDocument = 1
    End If
    Exit Function
    
SafeExit:
    GetParagraphIndexInDocument = 0
End Function

' 현재 문단 텍스트를 잘라서(개행 제거) 미리보기로 반환
Private Function GetCurrentParagraphPreview(ByVal rng As Range, ByVal maxLen As Long) As String
    On Error GoTo SafeExit
    
    Dim t As String
    t = rng.Paragraphs(1).Range.Text
    t = NormalizeInlineText(t)
    If maxLen > 0 And Len(t) > maxLen Then
        t = Left$(t, maxLen) & "…"
    End If
    GetCurrentParagraphPreview = t
    Exit Function
    
SafeExit:
    GetCurrentParagraphPreview = ""
End Function

' 커서 위치의 단어(대략). 빈 값이면 "".
Private Function GetWordAtCursor(ByVal rng As Range) As String
    On Error GoTo SafeExit
    
    Dim w As Range
    Set w = rng.Duplicate
    
    ' 선택 영역이 있으면 첫 단어 기준, 커서만 있으면 word 확장
    If Selection.Type = wdSelectionIP Then
        w.Expand wdWord
    Else
        w.SetRange rng.Start, rng.Start
        w.Expand wdWord
    End If
    
    Dim s As String
    s = NormalizeInlineText(w.Text)
    s = Trim$(s)
    
    GetWordAtCursor = s
    Exit Function
    
SafeExit:
    GetWordAtCursor = ""
End Function


' 해당 Range가 포함되는 북마크 이름들을 최대 maxCount개까지 반환(쉼표로 구분)
Private Function GetBookmarkNamesAtRange(ByVal doc As Document, ByVal rng As Range, ByVal maxCount As Long) As String
    On Error GoTo SafeExit
    
    Dim names As String
    names = ""
    
    Dim bm As Bookmark
    Dim count As Long
    count = 0
    
    For Each bm In doc.Bookmarks
        ' 커서가 북마크 범위 내부에 있는지
        If rng.Start >= bm.Range.Start And rng.Start <= bm.Range.End Then
            count = count + 1
            If names <> "" Then names = names & ", "
            names = names & bm.Name
            If maxCount > 0 And count >= maxCount Then
                names = names & " …"
                Exit For
            End If
        End If
    Next bm
    
    GetBookmarkNamesAtRange = names
    Exit Function
    
SafeExit:
    GetBookmarkNamesAtRange = ""
End Function

' 해당 Range에 걸린 하이퍼링크 정보를 최대 maxCount개까지 반환
Private Function GetHyperlinkInfoAtRange(ByVal rng As Range, ByVal maxCount As Long) As String
    On Error GoTo SafeExit
    
    Dim info As String
    info = ""
    
    Dim hr As Range
    Set hr = rng.Duplicate
    
    ' 커서만 있는 경우: 주변 문자 1개 정도 포함해서 Hyperlinks를 잡을 확률을 높임
    If Selection.Type = wdSelectionIP Then
        On Error Resume Next
        If hr.Start > 0 Then hr.SetRange hr.Start - 1, hr.Start + 1
        On Error GoTo SafeExit
    End If
    
    Dim h As Hyperlink
    Dim c As Long
    c = 0
    
    If hr.Hyperlinks.Count > 0 Then
        For Each h In hr.Hyperlinks
            c = c + 1
            info = info & "- 표시: " & SafeInline(h.TextToDisplay) & vbCrLf
            info = info & "  Address: " & SafeInline(h.Address) & vbCrLf
            info = info & "  SubAddress: " & SafeInline(h.SubAddress) & vbCrLf
            If maxCount > 0 And c >= maxCount Then
                info = info & "  (… 더 있음)" & vbCrLf
                Exit For
            End If
        Next h
    End If
    
    ' 하이퍼링크가 Hyperlinks로 안 잡히는 경우: HYPERLINK 필드도 체크
    If info = "" Then
        Dim f As Field
        For Each f In rng.Fields
            If InStr(1, f.Code.Text, "HYPERLINK", vbTextCompare) > 0 Then
                info = "- (필드) " & NormalizeInlineText(f.Result.Text)
                Exit For
            End If
        Next f
    End If
    
    GetHyperlinkInfoAtRange = Trim$(info)
    Exit Function
    
SafeExit:
    GetHyperlinkInfoAtRange = ""
End Function

' 개행/탭 등을 한 줄 텍스트로 정리
Private Function NormalizeInlineText(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW$(160), " ") ' NBSP
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeInlineText = Trim$(s)
End Function

Private Function SafeInline(ByVal s As String) As String
    If s = "" Then
        SafeInline = "(없음)"
    Else
        SafeInline = NormalizeInlineText(s)
    End If
End Function
