Option Explicit

' 주의
' - Hyperlink.Follow의 AddHistory 인자는 문서상 "reserved for future use"로,
'   WebGoBack 히스토리를 쌓아주지 않는 경우가 있습니다.
' - 따라서 "돌아가기/앞으로가기"는 임시 북마크 스택으로 보장합니다.
' - 스택 구현은 `bookmarkStack.vba` 모듈로 분리되어 있습니다.
'
' Alt+→ 동작 확장
' ----------------------
' - 커서 위치에 하이퍼링크가 있으면 하이퍼링크 목적지로 이동(Follow)
' - 하이퍼링크가 없으면 우리가 저장한 앞 위치로 이동(Forward stack)

' Alt+← 동작 확장
' ----------------------
' - 우리가 저장한 뒤 위치로 이동(Back stack)
'
' 등록 예:
'   Call RegisterHotkey("NavigateForward", wdKeyAlt, wdKeyRight)

' Alt+→ 대체
Public Sub NavigateForward()
    On Error GoTo Fallback
    
    Dim rng As Range
    Set rng = Selection.Range.Duplicate
    
    ' 하이퍼링크 또는 클릭 가능한 Field를 따라갑니다.
    If TryFollowLinkOrFieldAt(rng) Then Exit Sub
    
    ' 우리가 저장한 위치가 있으면 그곳으로 이동합니다.
    If TryPopForwardLocation() Then Exit Sub
    
Fallback:
    On Error Resume Next
    Call Application.Run("WebGoForward")
End Sub

' Alt+← 대체
Public Sub NavigateBackward()
    On Error GoTo Fallback
    
    ' 우리가 저장한 위치가 있으면 그곳으로 이동합니다.
    If TryPopBackLocation() Then Exit Sub
    
Fallback:
    On Error Resume Next
    Call Application.Run("WebGoBack")
End Sub

' 커서 위치에서 Hyperlink 또는 클릭 가능한 Field(REF/HYPERLINK)를 따라갑니다.
' 성공 시 True.
Private Function TryFollowLinkOrFieldAt(ByVal rng As Range) As Boolean
    On Error GoTo SafeExit
    
    If rng Is Nothing Then GoTo SafeExit
    
    Dim link As Hyperlink
    Set link = GetFirstHyperlinkAt(rng)
    
    If Not link Is Nothing Then
        Call PushBackwardBookmark(rng)
        Call ClearNavForwardStack(rng.Document)
        Call link.Follow(NewWindow:=False)
        TryFollowLinkOrFieldAt = True
        Exit Function
    End If
    
    Dim fld As Field
    Set fld = GetFirstClickableFieldAt(rng)
    If fld Is Nothing Then GoTo SafeExit
    
    Dim bookmarkName As String
    bookmarkName = GetBookmarkNameFromClickableField(fld)
    
    If bookmarkName <> "" Then
        Call PushBackwardBookmark(rng)
        Call ClearNavForwardStack(rng.Document)
        Call GoToBookmarkInDocument(rng.Document, bookmarkName)
        TryFollowLinkOrFieldAt = True
        Exit Function
    End If
    
    Dim fieldLink As Hyperlink
    Set fieldLink = GetHyperlinkFromFieldResult(fld)
    If Not fieldLink Is Nothing Then
        Call PushBackwardBookmark(rng)
        Call ClearNavForwardStack(rng.Document)
        Call fieldLink.Follow(NewWindow:=False)
        TryFollowLinkOrFieldAt = True
        Exit Function
    End If
    
SafeExit:
    TryFollowLinkOrFieldAt = False
End Function

' 커서/선택 영역 주변에서 하이퍼링크 1개를 찾아 반환합니다. 없으면 Nothing.
Private Function GetFirstHyperlinkAt(ByVal rng As Range) As Hyperlink
    On Error GoTo SafeExit
    
    If rng Is Nothing Then GoTo SafeExit
    If rng.Start <> rng.End Then GoTo SafeExit
    
    ' 현재 Range에 하이퍼링크가 직접 걸린 경우
    If rng.Hyperlinks.Count > 0 Then
        Set GetFirstHyperlinkAt = rng.Hyperlinks(1)
        Exit Function
    End If
    
SafeExit:
    Set GetFirstHyperlinkAt = Nothing
End Function

' 커서 위치 주변에서 클릭 가능한 Field를 찾습니다. 없으면 Nothing.
Private Function GetFirstClickableFieldAt(ByVal rng As Range) As Field
    On Error GoTo SafeExit
    
    If rng Is Nothing Then GoTo SafeExit
    If rng.Start <> rng.End Then GoTo SafeExit
    
    Dim f As Field
    
    Dim i As Long
    For i = 0 To 7
        Dim probe As Range
        Set probe = ExpandRange(rng, i)
        Set f = GetFirstClickableFieldInRange(probe)
        If Not f Is Nothing Then
            Set GetFirstClickableFieldAt = f
            Exit Function
        End If
    Next i
    
SafeExit:
    Set GetFirstClickableFieldAt = Nothing
End Function

' Range 안에서 클릭 가능한 Field 1개를 반환합니다.
Private Function GetFirstClickableFieldInRange(ByVal rng As Range) As Field
    On Error GoTo SafeExit
    
    If rng Is Nothing Then GoTo SafeExit
    If rng.Fields.Count = 0 Then GoTo SafeExit
    
    Dim i As Long
    For i = 1 To rng.Fields.Count
        If IsClickableField(rng.Fields(i)) Then
            Set GetFirstClickableFieldInRange = rng.Fields(i)
            Exit Function
        End If
    Next i
    
SafeExit:
    Set GetFirstClickableFieldInRange = Nothing
End Function

Private Function IsClickableField(ByVal fld As Field) As Boolean
    On Error GoTo SafeExit
    
    If fld Is Nothing Then GoTo SafeExit
    
    If fld.Type = wdFieldHyperlink Then
        If GetBookmarkNameFromClickableField(fld) <> "" Then
            IsClickableField = True
            Exit Function
        End If
        
        If Not GetHyperlinkFromFieldResult(fld) Is Nothing Then
            IsClickableField = True
            Exit Function
        End If
    End If
    
    If fld.Type = wdFieldRef Then
        If GetBookmarkNameFromClickableField(fld) <> "" Then
            IsClickableField = True
            Exit Function
        End If
    End If
    
SafeExit:
    IsClickableField = False
End Function

' 클릭 가능한 Field에서 (문서 내부) 북마크 이름을 추출합니다. 없으면 "".
Private Function GetBookmarkNameFromClickableField(ByVal fld As Field) As String
    On Error GoTo SafeExit
    
    If fld Is Nothing Then GoTo SafeExit
    
    Dim codeText As String
    codeText = fld.Code.Text
    
    If fld.Type = wdFieldRef Then
        GetBookmarkNameFromClickableField = ParseBookmarkNameFromRefField(codeText)
        Exit Function
    End If
    
    If fld.Type = wdFieldHyperlink Then
        GetBookmarkNameFromClickableField = _
            ParseBookmarkNameFromHyperlinkField(codeText)
        Exit Function
    End If
    
SafeExit:
    GetBookmarkNameFromClickableField = ""
End Function

Private Function ParseBookmarkNameFromRefField(ByVal codeText As String) As String
    On Error GoTo SafeExit
    
    Dim s As String
    s = Trim$(codeText)
    
    If Len(s) < 3 Then GoTo SafeExit
    If LCase$(Left$(s, 3)) <> "ref" Then GoTo SafeExit
    
    Dim rest As String
    rest = Trim$(Mid$(s, 4))
    
    Dim token As String
    token = ReadFirstFieldArgToken(rest)
    ParseBookmarkNameFromRefField = token
    Exit Function
    
SafeExit:
    ParseBookmarkNameFromRefField = ""
End Function

Private Function ParseBookmarkNameFromHyperlinkField(ByVal codeText As String) As String
    On Error GoTo SafeExit
    
    Dim s As String
    s = Trim$(codeText)
    
    If Len(s) < 9 Then GoTo SafeExit
    If LCase$(Left$(s, 9)) <> "hyperlink" Then GoTo SafeExit
    
    Dim p As Long
    p = InStr(1, s, "\l", vbTextCompare)
    
    Dim rest As String
    If p > 0 Then
        rest = Trim$(Mid$(s, p + 2))
        ParseBookmarkNameFromHyperlinkField = ReadFirstFieldArgToken(rest)
        Exit Function
    End If
    
    rest = Trim$(Mid$(s, 10))
    Dim firstArg As String
    firstArg = ReadFirstFieldArgToken(rest)
    
    If Left$(firstArg, 1) = "#" Then
        ParseBookmarkNameFromHyperlinkField = Mid$(firstArg, 2)
        Exit Function
    End If
    
SafeExit:
    ParseBookmarkNameFromHyperlinkField = ""
End Function

' rest에서 첫 인자 토큰(따옴표/공백/백슬래시 구분)을 읽습니다.
Private Function ReadFirstFieldArgToken(ByVal rest As String) As String
    On Error GoTo SafeExit
    
    Dim s As String
    s = Trim$(rest)
    If s = "" Then GoTo SafeExit
    
    If Left$(s, 1) = """" Then
        Dim q As Long
        q = FindClosingQuotePos(s, 2)
        If q <= 2 Then GoTo SafeExit
        ReadFirstFieldArgToken = Replace$(Mid$(s, 2, q - 2), """""", """")
        Exit Function
    End If
    
    Dim pSpace As Long
    Dim pSlash As Long
    pSpace = InStr(1, s, " ")
    pSlash = InStr(1, s, "\")
    
    Dim pEnd As Long
    pEnd = 0
    If pSpace > 0 Then pEnd = pSpace
    If pSlash > 0 Then
        If pEnd = 0 Or pSlash < pEnd Then pEnd = pSlash
    End If
    
    If pEnd = 0 Then
        ReadFirstFieldArgToken = s
        Exit Function
    End If
    
    ReadFirstFieldArgToken = Trim$(Left$(s, pEnd - 1))
    Exit Function
    
SafeExit:
    ReadFirstFieldArgToken = ""
End Function

' s(startPos)부터 닫는 따옴표 위치를 찾습니다. 없으면 0.
' - """"(연속 2개)은 이스케이프 처리로 간주하고 건너뜁니다.
Private Function FindClosingQuotePos(ByVal s As String, ByVal startPos As Long) As Long
    On Error GoTo SafeExit
    
    Dim i As Long
    Dim n As Long
    n = Len(s)
    
    If startPos < 1 Then startPos = 1
    
    i = startPos
    Do While i <= n
        If Mid$(s, i, 1) = """" Then
            If i < n And Mid$(s, i + 1, 1) = """" Then
                i = i + 2
            Else
                FindClosingQuotePos = i
                Exit Function
            End If
        Else
            i = i + 1
        End If
    Loop
    
SafeExit:
    FindClosingQuotePos = 0
End Function

Private Sub GoToBookmarkInDocument(ByVal doc As Document, ByVal bookmarkName As String)
    On Error GoTo SafeExit
    
    If doc Is Nothing Then GoTo SafeExit
    If bookmarkName = "" Then GoTo SafeExit
    If Not doc.Bookmarks.Exists(bookmarkName) Then GoTo SafeExit
    
    Call doc.Activate
    Call Selection.GoTo(What:=wdGoToBookmark, Name:=bookmarkName)
    
SafeExit:
End Sub

Private Function GetHyperlinkFromFieldResult(ByVal fld As Field) As Hyperlink
    On Error GoTo SafeExit
    
    If fld Is Nothing Then GoTo SafeExit
    If fld.Result Is Nothing Then GoTo SafeExit
    If fld.Result.Hyperlinks.Count = 0 Then GoTo SafeExit
    
    Set GetHyperlinkFromFieldResult = fld.Result.Hyperlinks(1)
    Exit Function
    
SafeExit:
    Set GetHyperlinkFromFieldResult = Nothing
End Function

' 커서 주변 nChars만큼 확장한 Range를 만들어 Field 탐지에 사용합니다.
Private Function ExpandRange(ByVal rng As Range, ByVal nChars As Long) As Range
    On Error GoTo Fallback
    
    If nChars <= 0 Then
        Set ExpandRange = rng.Duplicate
        Exit Function
    End If
    
    Dim doc As Document
    Set doc = rng.Document
    
    Dim probe As Range
    Set probe = rng.Duplicate
    
    If probe.Start > nChars Then
        probe.Start = probe.Start - nChars
    Else
        probe.Start = 0
    End If
    
    Dim docEnd As Long
    docEnd = doc.Content.End
    
    If probe.End + nChars < docEnd Then
        probe.End = probe.End + nChars
    Else
        probe.End = docEnd
    End If
    
    Set ExpandRange = probe
    Exit Function
    
Fallback:
    Set ExpandRange = rng
End Function

