Option Explicit

' =============================================================================
' Module: cursorTracker
' Purpose:
'   - 문서별 "현재/이전" 커서 위치 및 관련 정보를 메모리에 유지합니다.
'   - 이 모듈은 메모리 상태만 관리하며, 영속 저장/로드는 별도 모듈이 담당합니다.
'
' Public API:
'   - InitializeCursorMoveTracking()
'   - ToggleCursorMoveTracking()
'   - TrackCursorMove(targetRange As Range)
'   - CursorMemory_TryGetCurrent(doc As Document, outInfo As CursorMoveInfo) _
'       As Boolean
'   - CursorMemory_TryGetPrevious(doc As Document, outInfo As CursorMoveInfo) _
'       As Boolean
'   - a_ShowCursorLocationInfo()  ' 디버깅용 매크로
'
' Related:
'   - myVba/_utils/cursorMoveStack.vba: 커서 이동 히스토리(CustomXMLParts) 저장/로드
'   - myVba/_utils/CursorMoveInfo.cls: 커서 위치 정보 컨테이너
' =============================================================================

Private Const CURSOR_INFO_INDEX_PREV As Long = 0
Private Const CURSOR_INFO_INDEX_CURR As Long = 1

Private gIsCursorMoveTrackingEnabled As Boolean

' docKey -> Variant(Array(prevInfo, currInfo))
Private gPrevCurrByDoc As Object ' Scripting.Dictionary (late-bound)

Private Sub EnsurePrevCurrDict()
    If gPrevCurrByDoc Is Nothing Then
        Set gPrevCurrByDoc = CreateObject("Scripting.Dictionary")
    End If
End Sub

Private Function GetDocumentKey(ByVal doc As Document) As String
    On Error GoTo Fallback
    If Not doc Is Nothing Then
        If doc.FullName <> "" Then
            GetDocumentKey = doc.FullName
            Exit Function
        End If
        GetDocumentKey = doc.Name & "#" & Hex$(ObjPtr(doc))
        Exit Function
    End If
Fallback:
    GetDocumentKey = "(unknown)"
End Function

Private Sub UpdatePrevCurrInMemory(ByVal doc As Document, ByVal newInfo As CursorMoveInfo)
    On Error GoTo SafeExit
    If doc Is Nothing Then Exit Sub
    If newInfo Is Nothing Then Exit Sub

    Call EnsurePrevCurrDict()
    Dim k As String
    k = GetDocumentKey(doc)

    Dim v As Variant
    Dim prevInfo As CursorMoveInfo
    Dim currInfo As CursorMoveInfo

    If gPrevCurrByDoc.Exists(k) Then
        v = gPrevCurrByDoc(k)
        If IsArray(v) Then
            If Not IsEmpty(v(CURSOR_INFO_INDEX_CURR)) Then
                Set currInfo = v(CURSOR_INFO_INDEX_CURR)
            End If
            If Not IsEmpty(v(CURSOR_INFO_INDEX_PREV)) Then
                Set prevInfo = v(CURSOR_INFO_INDEX_PREV)
            End If
        End If
    End If

    ' shift: curr -> prev, new -> curr
    Set prevInfo = currInfo
    Set currInfo = newInfo

    v = Array(Empty, Empty)
    If Not prevInfo Is Nothing Then Set v(CURSOR_INFO_INDEX_PREV) = prevInfo
    If Not currInfo Is Nothing Then Set v(CURSOR_INFO_INDEX_CURR) = currInfo
    gPrevCurrByDoc(k) = v

SafeExit:
End Sub

Public Function CursorMemory_TryGetCurrent( _
    ByVal doc As Document, _
    ByRef outInfo As CursorMoveInfo _
) As Boolean
    On Error GoTo SafeExit
    Call EnsurePrevCurrDict()
    Dim k As String
    k = GetDocumentKey(doc)
    If Not gPrevCurrByDoc.Exists(k) Then GoTo SafeExit

    Dim v As Variant
    v = gPrevCurrByDoc(k)
    If Not IsArray(v) Then GoTo SafeExit
    If IsEmpty(v(CURSOR_INFO_INDEX_CURR)) Then GoTo SafeExit

    Set outInfo = v(CURSOR_INFO_INDEX_CURR)
    CursorMemory_TryGetCurrent = True
    Exit Function

SafeExit:
    CursorMemory_TryGetCurrent = False
End Function

Public Function CursorMemory_TryGetPrevious( _
    ByVal doc As Document, _
    ByRef outInfo As CursorMoveInfo _
) As Boolean
    On Error GoTo SafeExit
    Call EnsurePrevCurrDict()
    Dim k As String
    k = GetDocumentKey(doc)
    If Not gPrevCurrByDoc.Exists(k) Then GoTo SafeExit

    Dim v As Variant
    v = gPrevCurrByDoc(k)
    If Not IsArray(v) Then GoTo SafeExit
    If IsEmpty(v(CURSOR_INFO_INDEX_PREV)) Then GoTo SafeExit

    Set outInfo = v(CURSOR_INFO_INDEX_PREV)
    CursorMemory_TryGetPrevious = True
    Exit Function

SafeExit:
    CursorMemory_TryGetPrevious = False
End Function

' (MARK) (디버깅용) 현재 커서 위치 정보 알림
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
' - 매크로: a_ShowCursorLocationInfo 실행

' (MARK) 문서별 커서 이동 히스토리(CustomXMLParts) 저장
' ----------------------
' - 커서 이동(SelectionChange)마다 "현재/이전" 커서 정보는 메모리에만 유지합니다.
' - (커서 이동 히스토리(CustomXMLParts) 저장/로드는 cursorMoveStack.vba 에서 관리)
' - 이벤트 훅: clsAppEvents.cls 의 appWord_WindowSelectionChange 에서
'   TrackCursorMove(Sel.Range) 을 호출하도록 연결하면 동작합니다.

Public Sub a_ShowCursorLocationInfo()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim msg As String
    
    Set doc = ActiveDocument
    
    Dim lastInfo As CursorMoveInfo
    ' 메모리(current)에서 가져오기
    If Not CursorMemory_TryGetCurrent(doc, lastInfo) Then
        VBA.MsgBox _
            "메모리에 저장된 커서 정보가 없습니다.", _
            vbInformation, _
            "커서 위치 정보"
        Exit Sub
    End If
    
    Dim selRng As Range
    On Error Resume Next
    Set selRng = doc.Range(lastInfo.Position, lastInfo.Position)
    On Error GoTo ErrorHandler
    If selRng Is Nothing Then
        VBA.MsgBox _
            "최근 커서 히스토리의 위치(pos)를 Range로 변환할 수 없습니다.", _
            vbInformation, _
            "커서 위치 정보"
        Exit Sub
    End If
    
    ' ===== 페이지 =====
    Dim pageNo As Long
    pageNo = lastInfo.PageNo
    
    ' ===== 영역 제목(탐색창:목차에 보이는 제목) =====
    Dim headingTitle As String
    headingTitle = GetCurrentHeadingTitle(selRng, 200)
    
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
    msg = msg & "영역 제목: " & SafeInline(headingTitle) & vbCrLf
    msg = msg & "현재 단어: " & SafeInline(wordText) & vbCrLf
    msg = msg & "북마크: " & SafeInline(bmNames) & vbCrLf
    msg = msg & "하이퍼링크: " & SafeInline(subAddress) & vbCrLf
    msg = msg & vbCrLf & "문단 미리보기:" & vbCrLf & paraText
    
    ' 기본 MsgBox (원하면 showMsg로 교체 가능)
    VBA.MsgBox msg, vbInformation, "커서 위치 정보"
    Exit Sub
    
ErrorHandler:
    VBA.MsgBox "오류: " & Err.Description, vbCritical, "커서 위치 정보"
End Sub

' 초기화
Public Sub InitializeCursorMoveTracking()
    gIsCursorMoveTrackingEnabled = True
    Call EnsurePrevCurrDict()
End Sub

' Alt+R 토글용 매크로
Public Sub ToggleCursorMoveTracking()
    gIsCursorMoveTrackingEnabled = Not gIsCursorMoveTrackingEnabled
    
    Dim stateLabel As String
    stateLabel = IIf(gIsCursorMoveTrackingEnabled, "ON", "OFF")
    
    VBA.MsgBox _
        "커서 이동 추적 상태: " & stateLabel, _
        vbInformation, _
        "Cursor Move Tracking"
End Sub

' (MARK) WindowSelectionChange 이벤트 핸들러
Public Sub TrackCursorMove(ByVal targetRange As Range)
    On Error GoTo SafeExit
    
    If Not gIsCursorMoveTrackingEnabled Then Exit Sub
    
    If targetRange Is Nothing Then Exit Sub
    If targetRange.Start <> targetRange.End Then Exit Sub
    
    Dim doc As Document
    Set doc = targetRange.Document
    
    Dim rng As Range
    Set rng = targetRange.Duplicate
    
    ' 너무 잦은 이벤트/중복 기록 방지 (문서+pos 기준)
    Static sLastDocKey As String
    Static sLastPos As Long
    
    Dim docKey As String
    docKey = GetDocumentKey(doc)
    
    If docKey = sLastDocKey And rng.Start = sLastPos Then Exit Sub
    
    Dim info As CursorMoveInfo
    Set info = BuildCursorMoveInfo(doc, rng)
    
    ' 현재/이전 커서 위치는 CustomXMLParts에 "바로 저장하지 않고" 메모리에만 유지
    Call UpdatePrevCurrInMemory(doc, info)
    
    sLastDocKey = docKey
    sLastPos = rng.Start
    
SafeExit:
End Sub

Private Function BuildCursorMoveInfo(ByVal doc As Document, ByVal rng As Range) As CursorMoveInfo
    On Error GoTo SafeExit
    Dim info As CursorMoveInfo
    Set info = New CursorMoveInfo
    
    info.Position = rng.Start
    info.PageNo = rng.Information(wdActiveEndPageNumber)
    info.WordText = GetWordAtCursor(rng)
    info.BookmarkNames = GetBookmarkNamesAtRange(doc, rng, 15)
    info.SubAddress = GetFirstHyperlinkSubAddressAtRange(rng)
    
    Set BuildCursorMoveInfo = info
    Exit Function
    
SafeExit:
    Set BuildCursorMoveInfo = Nothing
End Function

' 커서 주변 하이퍼링크의 SubAddress 1개(없으면 "")
Private Function GetFirstHyperlinkSubAddressAtRange(ByVal rng As Range) As String
    On Error GoTo SafeExit
    
    Dim hyperlinkRange As Range
    Set hyperlinkRange = rng.Duplicate
    
    ' 커서(IP) 기준으로 "단어" 범위로 확장해서 하이퍼링크를 잡을 확률을 높임
    hyperlinkRange.SetRange rng.Start, rng.Start
    hyperlinkRange.Expand wdWord
    
    If hyperlinkRange.Hyperlinks.Count > 0 Then
        Dim h As Hyperlink
        For Each h In hyperlinkRange.Hyperlinks
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
    
    Dim wordRange As Range
    Set wordRange = rng.Duplicate
    
    ' 커서(IP) 기준 단어
    wordRange.SetRange rng.Start, rng.Start
    wordRange.Expand wdWord
    
    Dim s As String
    s = NormalizeInlineText(wordRange.Text)
    s = Trim$(s)
    
    GetWordAtCursor = s
    Exit Function
    
SafeExit:
    GetWordAtCursor = ""
End Function


' 해당 Range가 포함되는 북마크 이름들을 최대 maxCount개까지 반환(쉼표로 구분)
Private Function GetBookmarkNamesAtRange( _
    ByVal doc As Document, _
    ByVal rng As Range, _
    ByVal maxCount As Long _
) As String
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
