Option Explicit

' 주의
' - Hyperlink.Follow의 AddHistory 인자는 문서상 "reserved for future use"로,
'   WebGoBack 히스토리를 쌓아주지 않는 경우가 있습니다.
' - 따라서 "돌아가기"는 별도 스택(커서 Range 정보)으로 보장합니다.
'   참고: [Hyperlink.Follow method (Word)]
'   https://learn.microsoft.com/en-us/office/vba/api/word.hyperlink.follow
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
'   Call RegisterHotkey("FollowHyperlinkOrGoForwardPos", wdKeyAlt, wdKeyRight)

' Nav stack item
' - Word.Range 객체를 저장합니다.
'
' 문서별 스택:
' - docKey -> stack
Private gNavBackByDoc As Object ' Scripting.Dictionary (late-bound)
Private gNavForwardByDoc As Object ' Scripting.Dictionary (late-bound)

Public Sub FollowHyperlinkOrGoForwardPos()
    On Error GoTo Fallback
    
    Dim rng As Range
    Set rng = Selection.Range
    
    Dim link As Hyperlink
    Set link = GetFirstHyperlinkAt(rng)    
    
    If Not link Is Nothing Then
        Call PushBackLocation(rng)
        Call ClearNavForwardStack(rng.Document)
        Call link.Follow(NewWindow:=False)
        Exit Sub
    Else
        If TryPopForwardLocation() Then Exit Sub
    End If
    
Fallback:
    On Error Resume Next
    Call Application.Run("WebGoForward")
End Sub

' 현재 위치를 Range 정보로 저장(스택 push)
Private Sub PushBackLocation(ByVal rng As Range)
    On Error GoTo SafeExit
    
    Dim doc As Document
    Set doc = rng.Document

    Dim s As stack
    Set s = GetBackStack(doc)
    
    Dim mark As Range
    Set mark = rng.Duplicate

    ' 선택 영역(Start~End)을 그대로 저장
    Call s.Push(mark)
    
SafeExit:
End Sub

' 현재 위치를 Range 정보로 저장(Forward stack push)
Private Sub PushForwardLocation(ByVal rng As Range)
    On Error GoTo SafeExit

    Dim doc As Document
    Set doc = rng.Document

    Dim s As stack
    Set s = GetForwardStack(doc)

    Dim mark As Range
    Set mark = rng.Duplicate

    ' 선택 영역(Start~End)을 그대로 저장
    Call s.Push(mark)

SafeExit:
End Sub

' Alt+← 대체: 우리가 저장한 위치가 있으면 그곳으로, 없으면 기본 WebGoBack
Public Sub GoBackwardPos()
    On Error GoTo Fallback
    
    If Not TryPopBackLocation() Then GoTo Fallback
    Exit Sub
    
Fallback:
    On Error Resume Next
    Call Application.Run("WebGoBack")
End Sub

' Alt+→ 대체(Forward): 우리가 저장한 위치가 있으면 그곳으로, 없으면 기본 WebGoForward
Public Sub GoForwardPos()
    On Error GoTo Fallback

    If TryPopForwardLocation() Then Exit Sub

Fallback:
    On Error Resume Next
    Call Application.Run("WebGoForward")
End Sub

' ======================
' Back stack (cursor Range)
' ======================

' 뒤로 스택 지우기(문서별)
Public Sub ClearNavBackStack(ByVal doc As Document)
    On Error GoTo SafeExit
    
    If doc Is Nothing Then Exit Sub
   
    Call EnsureNavDictionaries()
    
    Dim docKey As String
    docKey = GetDocumentKey(doc)

    If gNavBackByDoc.Exists(docKey) Then
        Call gNavBackByDoc.Remove(docKey)
    End If

    If gNavForwardByDoc.Exists(docKey) Then
        Call gNavForwardByDoc.Remove(docKey)
    End If
    
SafeExit:
End Sub

' 앞으로 스택 지우기(문서별)
Public Sub ClearNavForwardStack(ByVal doc As Document)
    On Error GoTo SafeExit

    If doc Is Nothing Then Exit Sub

    Call EnsureNavDictionaries()

    Dim docKey As String
    docKey = GetDocumentKey(doc)

    If gNavForwardByDoc.Exists(docKey) Then
        Call gNavForwardByDoc.Remove(docKey)
    End If

SafeExit:
End Sub

' 스택 pop 후 해당 Range로 이동. 성공하면 True
Private Function TryPopBackLocation() As Boolean
    On Error GoTo SafeExit

    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then GoTo SafeExit

    Dim backStack As stack
    Set backStack = GetBackStack(doc)
    If backStack.IsEmpty() Then GoTo SafeExit

    Dim target As Range
    Set target = backStack.Pop()
    If target Is Nothing Then GoTo SafeExit

    ' Back으로 이동하기 전, 현재 위치를 Forward stack에 저장
    Dim cur As Range
    On Error Resume Next
    Set cur = Selection.Range
    On Error GoTo SafeExit
    If Not cur Is Nothing Then
        Call PushForwardLocation(cur)
    End If

    Call target.Document.Activate
    Call target.Select
    
    TryPopBackLocation = True
    Exit Function
    
SafeExit:
    TryPopBackLocation = False
End Function

' 스택 pop 후 해당 Range로 이동(Forward). 성공하면 True
Private Function TryPopForwardLocation() As Boolean
    On Error GoTo SafeExit

    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then GoTo SafeExit

    Dim fwdStack As stack
    Set fwdStack = GetForwardStack(doc)
    If fwdStack.IsEmpty() Then GoTo SafeExit

    Dim target As Range
    Set target = fwdStack.Pop()
    If target Is Nothing Then GoTo SafeExit

    ' Forward로 이동하기 전, 현재 위치를 Back stack에 저장
    Dim cur As Range
    On Error Resume Next
    Set cur = Selection.Range
    On Error GoTo SafeExit
    If Not cur Is Nothing Then
        Call PushBackLocation(cur)
    End If

    Call target.Document.Activate
    Call target.Select

    TryPopForwardLocation = True
    Exit Function

SafeExit:
    TryPopForwardLocation = False
End Function

Private Function GetDocumentKey(ByVal doc As Document) As String
    On Error GoTo Fallback
    If doc Is Nothing Then GoTo Fallback
    
    If doc.FullName <> "" Then
        GetDocumentKey = doc.FullName
        Exit Function
    End If
    GetDocumentKey = doc.Name & "#" & Hex$(ObjPtr(doc))
    Exit Function
Fallback:
    GetDocumentKey = "(unknown)"
End Function

Private Sub EnsureNavDictionaries()
    If gNavBackByDoc Is Nothing Then
        Set gNavBackByDoc = CreateObject("Scripting.Dictionary")
    End If

    If gNavForwardByDoc Is Nothing Then
        Set gNavForwardByDoc = CreateObject("Scripting.Dictionary")
    End If
End Sub

Private Function GetBackStack(ByVal doc As Document) As stack
    Call EnsureNavDictionaries()

    Dim k As String
    k = GetDocumentKey(doc)

    If Not gNavBackByDoc.Exists(k) Then
        Dim s As stack
        Set s = New stack
        Call gNavBackByDoc.Add(k, s)
    End If

    Set GetBackStack = gNavBackByDoc(k)
End Function

Private Function GetForwardStack(ByVal doc As Document) As stack
    Call EnsureNavDictionaries()

    Dim k As String
    k = GetDocumentKey(doc)

    If Not gNavForwardByDoc.Exists(k) Then
        Dim s As stack
        Set s = New stack
        Call gNavForwardByDoc.Add(k, s)
    End If

    Set GetForwardStack = gNavForwardByDoc(k)
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

