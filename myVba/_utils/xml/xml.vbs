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

