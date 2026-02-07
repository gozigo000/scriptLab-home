Option Explicit

' ============================================================
' 모듈: xml
' 역할: XML/CustomXMLParts 관련 공통 헬퍼
' Public:
'   - EscapeXmlAttr: XML attribute escape
'   - EscapeXmlText: XML text escape
'   - CustomXml_FindPart: namespace/rootLocal로 CustomXMLPart 찾기
'   - CustomXml_DeleteExistingParts: namespace/rootLocal로 CustomXMLPart 삭제
'   - GetCustomXmlAttr: CustomXMLNode에서 attribute 읽기
'   - Xml_CreateDomDocument: MSXML DOMDocument 생성/기본 설정
'   - Xml_BuildXmlDeclaration: XML 선언 생성
'   - Xml_BuildAttr: attribute 문자열 생성(escape 포함)
'   - Xml_BuildStartTag: 시작 태그 생성
'   - Xml_BuildEndTag: 종료 태그 생성
'   - Xml_BuildEmptyElement: self-closing element 생성
'   - Xml_BuildElementText: 텍스트 element 생성(escape 포함)
'   - Xml_BuildRootStart: namespaced root 시작 태그 생성
'   - Xml_BuildRootEnd: namespaced root 종료 태그 생성
'   - Xml_BuildElement: element 생성(기본 self-closing)
'   - Xml_AddAttr: element에 attribute 추가(escape 포함)
'   - Xml_AddChild: element에 child XML 추가(필요 시 self-closing -> open/close)
'   - Xml_SetText: element의 text 설정(escape 포함)
' ============================================================

' XML attribute value escape
' - &, ", ', <, > 순으로 치환
Public Function EscapeXmlAttr(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    EscapeXmlAttr = s
End Function

' XML text escape
' - &, <, > 순으로 치환
Public Function EscapeXmlText(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    EscapeXmlText = s
End Function

' namespace/rootLocal 기준으로 파트 1개 찾기
' - SelectByNamespace 우선, 실패 시 전체 순회 fallback
Public Function CustomXml_FindPart( _
    ByVal doc As Document, _
    ByVal ns As String, _
    ByVal rootLocal As String _
) As CustomXMLPart
    On Error GoTo Fallback

    If doc Is Nothing Then
        Set CustomXml_FindPart = Nothing
        Exit Function
    End If

    Dim parts As CustomXMLParts
    Set parts = doc.CustomXMLParts.SelectByNamespace(ns)
    If Not parts Is Nothing Then
        If parts.Count > 0 Then
            Set CustomXml_FindPart = parts(1)
            Exit Function
        End If
    End If

Fallback:
    On Error Resume Next

    Dim p As CustomXMLPart
    For Each p In doc.CustomXMLParts
        If InStr(1, p.XML, ns, vbTextCompare) > 0 And _
            InStr(1, p.XML, rootLocal, vbTextCompare) > 0 _
        Then
            Set CustomXml_FindPart = p
            Exit Function
        End If
    Next p

    Set CustomXml_FindPart = Nothing
End Function

' namespace/rootLocal 기준으로 기존 파트들 삭제
' - SelectByNamespace 우선, 실패 시 전체 순회 fallback
Public Sub CustomXml_DeleteExistingParts( _
    ByVal doc As Document, _
    ByVal ns As String, _
    ByVal rootLocal As String _
)
    On Error GoTo SafeExit

    If doc Is Nothing Then Exit Sub

    Dim parts As CustomXMLParts
    Set parts = doc.CustomXMLParts.SelectByNamespace(ns)
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
        If InStr(1, p.XML, ns, vbTextCompare) > 0 And _
            InStr(1, p.XML, rootLocal, vbTextCompare) > 0 _
        Then
            p.Delete
        End If
    Next p
End Sub

' CustomXMLNode에서 attribute 읽기
' - Word CustomXMLNode는 NodeName 대신 BaseName을 사용하는 케이스가 많음
Public Function GetCustomXmlAttr( _
    ByVal node As CustomXMLNode, _
    ByVal attrName As String _
) As String
    On Error GoTo SafeExit

    If node Is Nothing Then GoTo SafeExit
    If node.Attributes Is Nothing Then GoTo SafeExit

    Dim a As CustomXMLNode
    For Each a In node.Attributes
        Dim bn As String
        bn = ""

        On Error Resume Next
        bn = CStr(a.BaseName)
        On Error GoTo SafeExit

        If LCase$(bn) = LCase$(attrName) Then
            GetCustomXmlAttr = CStr(a.Text)
            Exit Function
        End If
    Next a

SafeExit:
    GetCustomXmlAttr = ""
End Function

' MSXML DOMDocument.6.0 생성 및 기본 설정
' - async/validateOnParse/resolveExternals 옵션을 고정해서 보안/성능 리스크를 줄입니다.
' - 생성 실패 시 Nothing 반환
Public Function Xml_CreateDomDocument() As Object
    On Error GoTo SafeExit

    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")

    dom.async = False
    dom.validateOnParse = False
    dom.resolveExternals = False

    Set Xml_CreateDomDocument = dom
    Exit Function

SafeExit:
    Set Xml_CreateDomDocument = Nothing
End Function


' ============================
' XML 문자열 빌더
' ============================

Public Function Xml_BuildXmlDeclaration( _
    Optional ByVal encoding As String = "UTF-8" _
) As String
    Xml_BuildXmlDeclaration = _
        "<?xml version=""1.0"" encoding=""" & EscapeXmlAttr(encoding) & """?>"
End Function

' attribute 문자열을 생성합니다. (앞에 공백 포함)
' 예: Xml_BuildAttr("id", "a") =>  id="a"
Public Function Xml_BuildAttr(ByVal name As String, ByVal value As String) As String
    Xml_BuildAttr = " " & name & "=""" & EscapeXmlAttr(value) & """"
End Function

Public Function Xml_BuildStartTag( _
    ByVal tagName As String, _
    Optional ByVal attrs As String = "" _
) As String
    Xml_BuildStartTag = "<" & tagName & attrs & ">"
End Function

Public Function Xml_BuildEndTag(ByVal tagName As String) As String
    Xml_BuildEndTag = "</" & tagName & ">"
End Function

Public Function Xml_BuildEmptyElement( _
    ByVal tagName As String, _
    Optional ByVal attrs As String = "" _
) As String
    Xml_BuildEmptyElement = "<" & tagName & attrs & " />"
End Function

Public Function Xml_BuildElementText( _
    ByVal tagName As String, _
    ByVal text As String _
) As String
    Xml_BuildElementText = _
        "<" & tagName & ">" & EscapeXmlText(text) & "</" & tagName & ">"
End Function

Public Function Xml_BuildRootStart( _
    ByVal prefix As String, _
    ByVal rootLocal As String, _
    ByVal ns As String, _
    Optional ByVal version As String = "", _
    Optional ByVal extraAttrs As String = "" _
) As String
    Dim tagName As String
    tagName = prefix & ":" & rootLocal

    Dim attrs As String
    attrs = ""
    attrs = attrs & " xmlns:" & prefix & "=""" & EscapeXmlAttr(ns) & """"
    If version <> "" Then attrs = attrs & Xml_BuildAttr("version", version)
    If extraAttrs <> "" Then attrs = attrs & extraAttrs

    Xml_BuildRootStart = Xml_BuildStartTag(tagName, attrs)
End Function

Public Function Xml_BuildRootEnd( _
    ByVal prefix As String, _
    ByVal rootLocal As String _
) As String
    Xml_BuildRootEnd = Xml_BuildEndTag(prefix & ":" & rootLocal)
End Function

' element 생성(기본: self-closing)
' - 예: Xml_BuildElement("move") => "<move />"
Public Function Xml_BuildElement(ByVal tagName As String) As String
    Xml_BuildElement = Xml_BuildEmptyElement(tagName, "")
End Function

' element에 attribute 추가 (escape 포함)
' - nodeXml: "<tag />" 또는 "<tag .../>" 형태
' - 반환: attribute가 추가된 새 XML 문자열
Public Function Xml_AddAttr( _
    ByVal nodeXml As String, _
    ByVal name As String, _
    ByVal value As String _
) As String
    Dim p As Long
    p = Xml__FindTagInsertPos(nodeXml)
    If p <= 0 Then
        Xml_AddAttr = nodeXml
        Exit Function
    End If

    Xml_AddAttr = Left$(nodeXml, p - 1) & _
        Xml_BuildAttr(name, value) & _
        Mid$(nodeXml, p)
End Function

' element에 child XML을 추가합니다.
' - nodeXml이 self-closing이면 open/close로 변환 후 child를 넣습니다.
Public Function Xml_AddChild( _
    ByVal nodeXml As String, _
    ByVal childXml As String _
) As String
    Dim tagName As String
    tagName = Xml_GetTagName(nodeXml)
    If tagName = "" Then
        Xml_AddChild = nodeXml
        Exit Function
    End If

    Dim closeTag As String
    closeTag = "</" & tagName & ">"

    If Right$(Trim$(nodeXml), 2) = "/>" Then
        Dim openTag As String
        openTag = Left$(Trim$(nodeXml), Len(Trim$(nodeXml)) - 2) & ">"
        Xml_AddChild = openTag & childXml & closeTag
        Exit Function
    End If

    Dim posClose As Long
    posClose = InStrRev(nodeXml, closeTag)
    If posClose <= 0 Then
        Xml_AddChild = nodeXml
        Exit Function
    End If

    Xml_AddChild = Left$(nodeXml, posClose - 1) & childXml & Mid$(nodeXml, posClose)
End Function

' element의 text를 설정합니다. (escape 포함)
' - 기존 child가 있어도 text로 교체(간단 구현)
Public Function Xml_SetText( _
    ByVal nodeXml As String, _
    ByVal text As String _
) As String
    Dim tagName As String
    tagName = Xml_GetTagName(nodeXml)
    If tagName = "" Then
        Xml_SetText = nodeXml
        Exit Function
    End If

    Dim openEnd As Long
    openEnd = InStr(1, nodeXml, ">")
    If openEnd <= 0 Then
        Xml_SetText = nodeXml
        Exit Function
    End If

    Dim openTag As String
    Dim closeTag As String
    openTag = Left$(nodeXml, openEnd)
    closeTag = "</" & tagName & ">"

    If Right$(Trim$(nodeXml), 2) = "/>" Then
        openTag = Left$(Trim$(nodeXml), Len(Trim$(nodeXml)) - 2) & ">"
    End If

    Xml_SetText = openTag & EscapeXmlText(text) & closeTag
End Function

' "<tag" 다음 위치(삽입 지점)를 찾습니다.
Private Function Xml__FindTagInsertPos(ByVal nodeXml As String) As Long
    On Error GoTo SafeExit

    If Left$(nodeXml, 1) <> "<" Then GoTo SafeExit

    Dim pSpace As Long
    Dim pGt As Long
    Dim pSlash As Long

    pSpace = InStr(1, nodeXml, " ")
    pGt = InStr(1, nodeXml, ">")
    pSlash = InStr(1, nodeXml, "/>")

    Dim p As Long
    p = 0
    If pSpace > 0 Then p = pSpace
    If p = 0 Or (pGt > 0 And pGt < p) Then p = pGt
    If p = 0 Or (pSlash > 0 And pSlash < p) Then p = pSlash

    Xml__FindTagInsertPos = p
    Exit Function

SafeExit:
    Xml__FindTagInsertPos = 0
End Function

' "<tag ...>"에서 tagName을 추출합니다. (prefix 포함 가능)
Private Function Xml_GetTagName(ByVal nodeXml As String) As String
    On Error GoTo SafeExit

    Dim s As String
    s = Trim$(nodeXml)
    If Left$(s, 1) <> "<" Then GoTo SafeExit
    If Left$(s, 2) = "</" Then GoTo SafeExit

    Dim i As Long
    For i = 2 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch = " " Or ch = ">" Or ch = "/" Then
            Xml_GetTagName = Mid$(s, 2, i - 2)
            Exit Function
        End If
    Next i

SafeExit:
    Xml_GetTagName = ""
End Function


