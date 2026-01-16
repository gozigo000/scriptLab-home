Option Explicit

' ======================
' Cursor History (CustomXMLParts)
' ======================

Private Const CURSOR_XML_NS As String = "urn:scriptlab:cursorhistory:v1"
Private Const CURSOR_XML_ROOT_LOCAL As String = "cursorHistory"
Private Const CURSOR_XML_VERSION As String = "1"
Private Const CURSOR_HISTORY_MAX As Long = 300

Private gCursorHistoryEnabled As Boolean

' (MARK) 커서 위치 정보 알림 매크로
' ----------------------
' 현재 커서(Selection) 기준으로 다음 정보를 알림창으로 표시합니다.
' - 페이지 번호
' - 영역 제목(탐색창:목차에 보이는 제목)
' - 문단 텍스트(앞부분)
' - 현재 단어(커서 위치 기준)
' - 북마크 포함 여부(해당 커서가 속한 북마크 이름들)
' - 하이퍼링크 SubAddress
'
' 사용:
' - 매크로: ShowCursorLocationInfo 실행

' (MARK) 문서별 커서 이동 히스토리(CustomXMLParts) 저장
' ----------------------
' - 커서 이동(SelectionChange)마다 다음 정보를 문서의 CustomXMLParts에 누적 저장합니다.
'   pageNo, word, bookmarkNames, subAddress, pos
' - 이벤트 훅: clsAppEvents.cls 의 appWord_WindowSelectionChange 에서
'   CursorHistory_LogSelectionChange Sel 을 호출하도록 연결하면 동작합니다.

Public Sub a_ShowCursorLocationInfo()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim msg As String
    
    Set doc = ActiveDocument
    
    Dim lastInfo As CursorMoveInfo
    If Not CursorHistory_TryGetLatestMoveInfo(doc, lastInfo) Then
        VBA.MsgBox "저장된 커서 히스토리가 없습니다.", vbInformation, "커서 위치 정보"
        Exit Sub
    End If
    
    Dim selRng As Range
    On Error Resume Next
    Set selRng = doc.Range(lastInfo.Position, lastInfo.Position)
    On Error GoTo ErrorHandler
    If selRng Is Nothing Then
        VBA.MsgBox "최근 커서 히스토리의 위치(pos)를 Range로 변환할 수 없습니다.", vbInformation, "커서 위치 정보"
        Exit Sub
    End If
    
    ' ===== 페이지 =====
    Dim pageNo As Long
    pageNo = lastInfo.PageNo
    
    ' ===== 영역 제목(탐색창:목차에 보이는 제목) =====
    Dim headingTitle As String
    headingTitle = GetNearestHeadingTitle(selRng, 200)
    
    ' ===== 문단 =====
    Dim paraText As String
    paraText = GetCurrentParagraphPreview(selRng, 120)
    
    ' ===== 단어 =====
    Dim wordText As String
    wordText = lastInfo.WordText
    
    ' ===== 북마크 =====
    Dim bmNames As String
    bmNames = lastInfo.BookmarkNames
    
    ' ===== 하이퍼링크(SubAddress) =====
    Dim subAddress As String
    subAddress = lastInfo.SubAddress
    
    msg = ""
    msg = msg & "페이지: " & pageNo & vbCrLf
    msg = msg & "영역 제목: " & IIf(headingTitle = "", "(없음)", headingTitle) & vbCrLf
    msg = msg & "현재 단어: " & IIf(wordText = "", "(없음)", """" & wordText & """") & vbCrLf
    msg = msg & "북마크: " & IIf(bmNames = "", "(없음)", bmNames) & vbCrLf
    msg = msg & "하이퍼링크: " & IIf(subAddress = "", "(없음)", """" & subAddress & """") & vbCrLf
    msg = msg & vbCrLf & "문단 미리보기:" & vbCrLf & paraText
    
    ' 기본 MsgBox (원하면 showMsg로 교체 가능)
    VBA.MsgBox msg, vbInformation, "커서 위치 정보"
    Exit Sub
    
ErrorHandler:
    VBA.MsgBox "오류: " & Err.Description, vbCritical, "커서 위치 정보"
End Sub

' CustomXMLParts에 저장된 가장 최근 move를 CursorMoveInfo로 파싱
Private Function CursorHistory_TryGetLatestMoveInfo(ByVal doc As Document, ByRef outInfo As CursorMoveInfo) As Boolean
    On Error GoTo SafeExit
    
    Dim part As CustomXMLPart
    Set part = FindCursorHistoryPart(doc)
    If part Is Nothing Then GoTo SafeExit
    
    Dim moves As CustomXMLNodes
    Set moves = part.SelectNodes("/*[local-name()='" & CURSOR_XML_ROOT_LOCAL & "']/*[local-name()='move']")
    If moves Is Nothing Then GoTo SafeExit
    If moves.Count <= 0 Then GoTo SafeExit
    
    Dim n As CustomXMLNode
    Set n = moves(moves.Count) ' 가장 최근(마지막)
    If n Is Nothing Then GoTo SafeExit
    
    Dim info As CursorMoveInfo
    Set info = New CursorMoveInfo
    
    info.Position = CLng(Val(GetCustomXmlAttr(n, "pos")))
    info.PageNo = CLng(Val(GetCustomXmlAttr(n, "page")))
    info.WordText = GetCustomXmlAttr(n, "word")
    info.BookmarkNames = GetCustomXmlAttr(n, "bookmarks")
    info.SubAddress = GetCustomXmlAttr(n, "subAddress")
    
    Set outInfo = info
    CursorHistory_TryGetLatestMoveInfo = True
    Exit Function
    
SafeExit:
    CursorHistory_TryGetLatestMoveInfo = False
End Function

Private Function GetCustomXmlAttr(ByVal node As CustomXMLNode, ByVal attrName As String) As String
    On Error GoTo SafeExit
    
    If node Is Nothing Then GoTo SafeExit
    If node.Attributes Is Nothing Then GoTo SafeExit
    
    Dim a As CustomXMLNode
    For Each a In node.Attributes
        Dim bn As String
        bn = ""
        On Error Resume Next
        bn = CStr(a.BaseName) ' Word CustomXMLNode는 NodeName이 없고 BaseName을 사용
        On Error GoTo SafeExit
        
        If LCase$(bn) = LCase$(attrName) Then
            GetCustomXmlAttr = CStr(a.Text)
            Exit Function
        End If
    Next a
    
SafeExit:
    GetCustomXmlAttr = ""
End Function

' 초기화
Public Sub InitializeCursorHistory()
    gCursorHistoryEnabled = True
End Sub

' Alt+R 토글용 매크로
Public Sub ToggleCursorHistoryLogging()
    gCursorHistoryEnabled = Not gCursorHistoryEnabled
    
    VBA.MsgBox "커서 히스토리 기록: " & IIf(gCursorHistoryEnabled, "ON", "OFF"), vbInformation, "Cursor History"
End Sub

' (이벤트용) SelectionChange에서 호출
Public Sub CursorHistory_LogSelectionChange(ByVal Sel As Selection)
    On Error GoTo SafeExit
    
    If Not gCursorHistoryEnabled Then Exit Sub
    
    If Sel Is Nothing Then Exit Sub
    If Sel.Type <> wdSelectionIP Then Exit Sub ' 커서(IP)일 때만 기록
    
    Dim doc As Document
    Set doc = Sel.Range.Document
    
    Dim rng As Range
    Set rng = Sel.Range.Duplicate
    
    ' 너무 잦은 이벤트/중복 기록 방지 (문서+pos 기준)
    Static sLastDocId As String
    Static sLastPos As Long
    
    Dim docId As String
    docId = SafeDocId(doc)
    
    If docId = sLastDocId And rng.Start = sLastPos Then Exit Sub
    
    Dim info As CursorMoveInfo
    Set info = BuildCursorMoveInfo(doc, rng)
    Call CursorHistory_AppendToDocument(doc, info, CURSOR_HISTORY_MAX)
    
    sLastDocId = docId
    sLastPos = rng.Start
    
SafeExit:
End Sub

Private Function BuildCursorMoveInfo(ByVal doc As Document, ByVal rng As Range) As CursorMoveInfo
    Dim info As CursorMoveInfo
    Set info = New CursorMoveInfo
    
    info.Position = rng.Start
    info.PageNo = Selection.Information(wdActiveEndPageNumber)
    info.WordText = GetWordAtCursor(rng)
    info.BookmarkNames = GetBookmarkNamesAtRange(doc, rng, 15)
    info.SubAddress = GetFirstHyperlinkSubAddressAtRange(rng)
    
    Set BuildCursorMoveInfo = info
End Function

Private Sub CursorHistory_AppendToDocument(ByVal doc As Document, ByVal info As CursorMoveInfo, ByVal maxHistory As Long)
    On Error GoTo SafeExit
    
    Dim part As CustomXMLPart
    Set part = EnsureCursorHistoryPart(doc)
    If part Is Nothing Then Exit Sub
    
    Dim moveXml As String
    moveXml = BuildMoveNodeXml(info)
    
    Dim root As CustomXMLNode
    Set root = part.SelectSingleNode("/*[local-name()='" & CURSOR_XML_ROOT_LOCAL & "']")
    If root Is Nothing Then
        ' 파트가 깨졌으면 재생성
        Call DeleteExistingCursorHistoryParts(doc)
        Set part = EnsureCursorHistoryPart(doc)
        If part Is Nothing Then Exit Sub
        Set root = part.SelectSingleNode("/*[local-name()='" & CURSOR_XML_ROOT_LOCAL & "']")
        If root Is Nothing Then Exit Sub
    End If
    
    root.AppendChildSubtree moveXml
    
    ' 오래된 기록 Trim
    If maxHistory > 0 Then
        Call TrimCursorHistory(part, maxHistory)
    End If
    
SafeExit:
    ' 저장 실패는 무시(보호 문서/권한/환경 차이 등)
End Sub

Private Sub TrimCursorHistory(ByVal part As CustomXMLPart, ByVal maxHistory As Long)
    On Error GoTo SafeExit
    
    Dim nodes As CustomXMLNodes
    Set nodes = part.SelectNodes("/*[local-name()='" & CURSOR_XML_ROOT_LOCAL & "']/*[local-name()='move']")
    If nodes Is Nothing Then Exit Sub
    
    Do While nodes.Count > maxHistory
        nodes(1).Delete ' 가장 오래된 것부터 삭제
        Set nodes = part.SelectNodes("/*[local-name()='" & CURSOR_XML_ROOT_LOCAL & "']/*[local-name()='move']")
        If nodes Is Nothing Then Exit Sub
    Loop
    
SafeExit:
End Sub

Private Function EnsureCursorHistoryPart(ByVal doc As Document) As CustomXMLPart
    On Error GoTo SafeExit
    
    Dim part As CustomXMLPart
    Set part = FindCursorHistoryPart(doc)
    If Not part Is Nothing Then
        Set EnsureCursorHistoryPart = part
        Exit Function
    End If
    
    Dim xml As String
    xml = BuildEmptyCursorHistoryXml()
    doc.CustomXMLParts.Add xml
    
    Set EnsureCursorHistoryPart = FindCursorHistoryPart(doc)
    Exit Function
    
SafeExit:
    Set EnsureCursorHistoryPart = Nothing
End Function

Private Function FindCursorHistoryPart(ByVal doc As Document) As CustomXMLPart
    On Error GoTo Fallback
    
    Dim parts As CustomXMLParts
    Set parts = doc.CustomXMLParts.SelectByNamespace(CURSOR_XML_NS)
    If Not parts Is Nothing Then
        If parts.Count > 0 Then
            Set FindCursorHistoryPart = parts(1)
            Exit Function
        End If
    End If
    
Fallback:
    On Error Resume Next
    Dim p As CustomXMLPart
    For Each p In doc.CustomXMLParts
        If InStr(1, p.XML, CURSOR_XML_NS, vbTextCompare) > 0 And InStr(1, p.XML, CURSOR_XML_ROOT_LOCAL, vbTextCompare) > 0 Then
            Set FindCursorHistoryPart = p
            Exit Function
        End If
    Next p
    Set FindCursorHistoryPart = Nothing
End Function

Private Sub DeleteExistingCursorHistoryParts(ByVal doc As Document)
    On Error GoTo SafeExit
    Dim parts As CustomXMLParts
    Set parts = doc.CustomXMLParts.SelectByNamespace(CURSOR_XML_NS)
    If Not parts Is Nothing Then
        Do While parts.Count > 0
            parts(1).Delete
        Loop
        Exit Sub
    End If
SafeExit:
    On Error Resume Next
    Dim p As CustomXMLPart
    For Each p In doc.CustomXMLParts
        If InStr(1, p.XML, CURSOR_XML_NS, vbTextCompare) > 0 And InStr(1, p.XML, CURSOR_XML_ROOT_LOCAL, vbTextCompare) > 0 Then
            p.Delete
        End If
    Next p
End Sub

Private Function BuildEmptyCursorHistoryXml() As String
    BuildEmptyCursorHistoryXml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
        "<sl:" & CURSOR_XML_ROOT_LOCAL & " xmlns:sl=""" & CURSOR_XML_NS & """ version=""" & CURSOR_XML_VERSION & """>" & _
        "<meta />" & _
        "</sl:" & CURSOR_XML_ROOT_LOCAL & ">"
End Function

Private Function BuildMoveNodeXml(ByVal info As CursorMoveInfo) As String
    BuildMoveNodeXml = "<move pos=""" & CStr(info.Position) & _
        """ page=""" & CStr(info.PageNo) & _
        """ word=""" & EscapeXmlAttr(info.WordText) & _
        """ bookmarks=""" & EscapeXmlAttr(info.BookmarkNames) & _
        """ subAddress=""" & EscapeXmlAttr(info.SubAddress) & _
        """ />"
End Function

Private Function EscapeXmlAttr(ByVal s As String) As String
    ' XML attribute value escape
    s = Replace(s, "&", "&amp;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, "'", "&apos;")
    EscapeXmlAttr = s
End Function

Private Function SafeDocId(ByVal doc As Document) As String
    On Error GoTo Fallback
    If doc Is Nothing Then GoTo Fallback
    If doc.FullName <> "" Then
        SafeDocId = doc.FullName
        Exit Function
    End If
Fallback:
    SafeDocId = "(unknown)"
End Function

' 커서 주변 하이퍼링크의 SubAddress 1개(없으면 "")
Private Function GetFirstHyperlinkSubAddressAtRange(ByVal rng As Range) As String
    On Error GoTo SafeExit
    
    Dim hr As Range
    Set hr = rng.Duplicate
    
    ' 커서(IP) 기준으로 "단어" 범위로 확장해서 하이퍼링크를 잡을 확률을 높임
    hr.SetRange rng.Start, rng.Start
    hr.Expand wdWord
    
    If hr.Hyperlinks.Count > 0 Then
        Dim h As Hyperlink
        For Each h In hr.Hyperlinks
            GetFirstHyperlinkSubAddressAtRange = CStr(h.SubAddress)
            Exit Function
        Next h
    End If
    
SafeExit:
    GetFirstHyperlinkSubAddressAtRange = ""
End Function

' ======================
' Helpers
' ======================

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
    
    ' 커서(IP) 기준 단어
    w.SetRange rng.Start, rng.Start
    w.Expand wdWord
    
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
