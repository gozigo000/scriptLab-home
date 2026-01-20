Option Explicit

' ======================
' Cursor Tracker (In-Memory current/previous)
' ======================
' 
' - 현재/이전 커서 위치 및 관련 정보를 메모리에 유지
' - cursorTracker.vba는 현재/이전 커서 위치 정보만 관리하며,
'   cursorMoveStack.vba는 CustomXMLParts에 저장/로드를 관리.
'
' 사용 예:
' - 이벤트에서: Call WatchCursorMove(Sel.Range)
' - 필요할 때: Call CursorMemory_TryGetCurrent(doc, outInfo)
' - 필요할 때: Call CursorMemory_TryGetPrevious(doc, outInfo)

Private gCursorHistoryEnabled As Boolean

' docKey -> Variant(Array(prevInfo, currInfo))
Private gPrevCurrByDoc As Object ' Scripting.Dictionary (late-bound)

Private Sub EnsurePrevCurrDict()
    If gPrevCurrByDoc Is Nothing Then
        Set gPrevCurrByDoc = CreateObject("Scripting.Dictionary")
    End If
End Sub

Private Function DocKey(ByVal doc As Document) As String
    On Error GoTo Fallback
    If Not doc Is Nothing Then
        If doc.FullName <> "" Then
            DocKey = doc.FullName
            Exit Function
        End If
        DocKey = doc.Name & "#" & Hex$(ObjPtr(doc))
        Exit Function
    End If
Fallback:
    DocKey = "(unknown)"
End Function

Private Sub UpdatePrevCurrInMemory(ByVal doc As Document, ByVal newInfo As CursorMoveInfo)
    On Error GoTo SafeExit
    If doc Is Nothing Then Exit Sub
    If newInfo Is Nothing Then Exit Sub

    Call EnsurePrevCurrDict
    Dim k As String
    k = DocKey(doc)

    Dim v As Variant
    Dim prevInfo As CursorMoveInfo
    Dim currInfo As CursorMoveInfo

    If gPrevCurrByDoc.Exists(k) Then
        v = gPrevCurrByDoc(k)
        If IsArray(v) Then
            If (Not IsEmpty(v(1))) Then Set currInfo = v(1)
            If (Not IsEmpty(v(0))) Then Set prevInfo = v(0)
        End If
    End If

    ' shift: curr -> prev, new -> curr
    Set prevInfo = currInfo
    Set currInfo = newInfo

    v = Array(Empty, Empty)
    If Not prevInfo Is Nothing Then Set v(0) = prevInfo
    If Not currInfo Is Nothing Then Set v(1) = currInfo
    gPrevCurrByDoc(k) = v

SafeExit:
End Sub

Public Function CursorMemory_TryGetCurrent(ByVal doc As Document, ByRef outInfo As CursorMoveInfo) As Boolean
    On Error GoTo SafeExit
    Call EnsurePrevCurrDict
    Dim k As String
    k = DocKey(doc)
    If Not gPrevCurrByDoc.Exists(k) Then GoTo SafeExit

    Dim v As Variant
    v = gPrevCurrByDoc(k)
    If Not IsArray(v) Then GoTo SafeExit
    If IsEmpty(v(1)) Then GoTo SafeExit

    Set outInfo = v(1)
    CursorMemory_TryGetCurrent = True
    Exit Function

SafeExit:
    CursorMemory_TryGetCurrent = False
End Function

Public Function CursorMemory_TryGetPrevious(ByVal doc As Document, ByRef outInfo As CursorMoveInfo) As Boolean
    On Error GoTo SafeExit
    Call EnsurePrevCurrDict
    Dim k As String
    k = DocKey(doc)
    If Not gPrevCurrByDoc.Exists(k) Then GoTo SafeExit

    Dim v As Variant
    v = gPrevCurrByDoc(k)
    If Not IsArray(v) Then GoTo SafeExit
    If IsEmpty(v(0)) Then GoTo SafeExit

    Set outInfo = v(0)
    CursorMemory_TryGetPrevious = True
    Exit Function

SafeExit:
    CursorMemory_TryGetPrevious = False
End Function

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
' - 커서 이동(SelectionChange)마다 "현재/이전" 커서 정보는 메모리에만 유지합니다.
' - (커서 히스토리(CustomXMLParts) 저장/로드는 cursorMoveStack.vba 에서 관리)
' - 이벤트 훅: clsAppEvents.cls 의 appWord_WindowSelectionChange 에서
'   WatchCursorMove Sel.Range 을 호출하도록 연결하면 동작합니다.

Public Sub a_ShowCursorLocationInfo()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim msg As String
    
    Set doc = ActiveDocument
    
    Dim lastInfo As CursorMoveInfo
    ' 메모리(current)에서 가져오기
    If Not CursorMemory_TryGetCurrent(doc, lastInfo) Then
        VBA.MsgBox "메모리에 저장된 커서 정보가 없습니다.", vbInformation, "커서 위치 정보"
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

' 초기화
Public Sub InitializeCursorHistory()
    gCursorHistoryEnabled = True
    Call EnsurePrevCurrDict
End Sub

' Alt+R 토글용 매크로
Public Sub ToggleCursorHistoryLogging()
    gCursorHistoryEnabled = Not gCursorHistoryEnabled
    
    VBA.MsgBox "커서 히스토리 기록: " & IIf(gCursorHistoryEnabled, "ON", "OFF"), vbInformation, "Cursor History"
End Sub

' (MARK) WindowSelectionChange 이벤트 핸들러
Public Sub WatchCursorMove(ByVal targetRange As Range)
    On Error GoTo SafeExit
    
    If Not gCursorHistoryEnabled Then Exit Sub
    
    If targetRange Is Nothing Then Exit Sub
    If targetRange.Start <> targetRange.End Then Exit Sub
    
    Dim doc As Document
    Set doc = targetRange.Document
    
    Dim rng As Range
    Set rng = targetRange.Duplicate
    
    ' 너무 잦은 이벤트/중복 기록 방지 (문서+pos 기준)
    Static sLastDocId As String
    Static sLastPos As Long
    
    Dim docId As String
    docId = SafeDocId(doc)
    
    If docId = sLastDocId And rng.Start = sLastPos Then Exit Sub
    
    Dim info As CursorMoveInfo
    Set info = BuildCursorMoveInfo(doc, rng)
    
    ' 현재/이전 커서 위치는 CustomXMLParts에 "바로 저장하지 않고" 메모리에만 유지
    Call UpdatePrevCurrInMemory(doc, info)
    
    sLastDocId = docId
    sLastPos = rng.Start
    
SafeExit:
End Sub

Private Function BuildCursorMoveInfo(ByVal doc As Document, ByVal rng As Range) As CursorMoveInfo
    Dim info As CursorMoveInfo
    Set info = New CursorMoveInfo
    
    info.Position = rng.Start
    info.PageNo = rng.Information(wdActiveEndPageNumber)
    info.WordText = GetWordAtCursor(rng)
    info.BookmarkNames = GetBookmarkNamesAtRange(doc, rng, 15)
    info.SubAddress = GetFirstHyperlinkSubAddressAtRange(rng)
    
    Set BuildCursorMoveInfo = info
End Function

Private Function SafeDocId(ByVal doc As Document) As String
    ' 기존 로직 유지(중복방지용 docId)
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
