Option Explicit

' ======================
' Cursor Move Stack (In-Memory + Flush to CustomXMLParts)
' ======================
'
' - SelectionChange 때마다 CustomXMLParts에 "바로" 저장하지 않고,
'   메모리에 히스토리를 누적(Push)한 뒤 필요할 때 Flush로 저장합니다.
' - CustomXMLParts 저장/로드 관련 코드는 이 파일(cursorMoveStack.vba)에만 둡니다.
'
' 사용 예:
' - 이벤트에서: Call CursorMoveStack_Push(Sel.Range.Document, info)
' - 필요할 때: Call CursorMoveStack_FlushActiveDocument
'

Private Const CURSOR_XML_NS As String = "urn:scriptlab:cursorhistory:v1"
Private Const CURSOR_XML_ROOT_LOCAL As String = "cursorHistory"
Private Const CURSOR_XML_VERSION As String = "1"
Private Const CURSOR_HISTORY_MAX As Long = 50

' docKey -> Collection(of CursorMoveInfo)
Private gStacks As Object ' Scripting.Dictionary (late-bound)

Private Sub EnsureStacks()
    If gStacks Is Nothing Then
        Set gStacks = CreateObject("Scripting.Dictionary")
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

Private Function EnsureStack(ByVal doc As Document) As Collection
    Call EnsureStacks
    Dim k As String
    k = DocKey(doc)
    If Not gStacks.Exists(k) Then
        Dim c As Collection
        Set c = New Collection
        gStacks.Add k, c
    End If
    Set EnsureStack = gStacks(k)
End Function

' 메모리에 move를 누적
Public Sub CursorMoveStack_Push(ByVal doc As Document, ByVal info As CursorMoveInfo, Optional ByVal maxHistory As Long = CURSOR_HISTORY_MAX)
    On Error GoTo SafeExit
    If doc Is Nothing Then Exit Sub
    If info Is Nothing Then Exit Sub

    Dim c As Collection
    Set c = EnsureStack(doc)
    c.Add info

    If maxHistory > 0 Then
        Do While c.Count > maxHistory
            c.Remove 1
        Loop
    End If

SafeExit:
End Sub

Public Sub CursorMoveStack_ClearMemory(ByVal doc As Document)
    On Error GoTo SafeExit
    Call EnsureStacks
    Dim k As String
    k = DocKey(doc)
    If gStacks.Exists(k) Then gStacks.Remove k
SafeExit:
End Sub

' 메모리의 히스토리를 문서(CustomXMLParts)에 "저장(Flush)" - 기존 파트는 덮어쓰기
Public Sub CursorMoveStack_FlushToDocument(ByVal doc As Document, Optional ByVal maxHistory As Long = CURSOR_HISTORY_MAX, Optional ByVal clearAfter As Boolean = False)
    On Error GoTo SafeExit
    If doc Is Nothing Then Exit Sub

    Dim c As Collection
    Set c = EnsureStack(doc)
    If c Is Nothing Then Exit Sub
    If c.Count <= 0 Then Exit Sub

    ' 기존 저장본은 제거하고 새로 작성(중복/증식 방지)
    Call CustomXml_DeleteExistingParts( _
        doc, _
        CURSOR_XML_NS, _
        CURSOR_XML_ROOT_LOCAL _
    )

    Dim part As CustomXMLPart
    Set part = EnsureCursorHistoryPart(doc)
    If part Is Nothing Then Exit Sub

    Dim root As CustomXMLNode
    Set root = part.SelectSingleNode("/*[local-name()='" & CURSOR_XML_ROOT_LOCAL & "']")
    If root Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To c.Count
        Dim info As CursorMoveInfo
        Set info = c(i)
        root.AppendChildSubtree BuildMoveNodeXml(info)
    Next i

    ' 최종 Trim (안전)
    If maxHistory > 0 Then
        Call TrimCursorHistory(part, maxHistory)
    End If

    If clearAfter Then
        Call CursorMoveStack_ClearMemory(doc)
    End If

SafeExit:
End Sub

' 편의 매크로: 현재 문서 Flush
Public Sub CursorMoveStack_FlushActiveDocument()
    On Error GoTo SafeExit
    Call CursorMoveStack_FlushToDocument(ActiveDocument, CURSOR_HISTORY_MAX, False)
SafeExit:
End Sub

' ----------------------
' CustomXMLParts: Load/Parse
' ----------------------

' CustomXMLParts에 저장된 가장 최근 move를 CursorMoveInfo로 파싱
Public Function CursorMoveStack_TryGetLatestMoveInfoFromDocument(ByVal doc As Document, ByRef outInfo As CursorMoveInfo) As Boolean
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
    CursorMoveStack_TryGetLatestMoveInfoFromDocument = True
    Exit Function

SafeExit:
    CursorMoveStack_TryGetLatestMoveInfoFromDocument = False
End Function

' ----------------------
' CustomXMLParts: Create/Delete/Trim
' ----------------------

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
    Set FindCursorHistoryPart = CustomXml_FindPart( _
        doc, _
        CURSOR_XML_NS, _
        CURSOR_XML_ROOT_LOCAL _
    )
End Function

Private Sub TrimCursorHistory(ByVal part As CustomXMLPart, ByVal maxHistory As Long)
    On Error GoTo SafeExit

    Dim nodes As CustomXMLNodes
    Set nodes = part.SelectNodes( _
        "/*[local-name()='" & CURSOR_XML_ROOT_LOCAL & "']" & _
        "/*[local-name()='move']" _
    )
    If nodes Is Nothing Then Exit Sub

    Do While nodes.Count > maxHistory
        nodes(1).Delete ' 가장 오래된 것부터 삭제
        Set nodes = part.SelectNodes( _
            "/*[local-name()='" & CURSOR_XML_ROOT_LOCAL & "']" & _
            "/*[local-name()='move']" _
        )
        If nodes Is Nothing Then Exit Sub
    Loop

SafeExit:
End Sub

Private Function BuildEmptyCursorHistoryXml() As String
    BuildEmptyCursorHistoryXml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
        "<sl:" & CURSOR_XML_ROOT_LOCAL & _
        " xmlns:sl=""" & CURSOR_XML_NS & & _
        """ version=""" & CURSOR_XML_VERSION & """>" & _
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
