Option Explicit

' Bookmark navigation stack
' -------------------------
' Word의 WebGoBack/WebGoForward 히스토리가 쌓이지 않는 케이스가 있어,
' "돌아가기/앞으로가기"를 임시 북마크 이름 스택으로 보장합니다.
'
' Public API
' - PushBackwardBookmark(rng): 현재 위치를 Back stack에 push
' - TryPopBackLocation(): Back stack pop 후 해당 위치로 이동(성공 True)
' - TryPopForwardLocation(): Forward stack pop 후 해당 위치로 이동(성공 True)
' - ClearNavBackStack(doc): 문서별 Back/Forward 스택 모두 삭제
' - ClearNavForwardStack(doc): 문서별 Forward 스택 삭제
'
' 의존
' - `myVba\_utils\dataStructure\Stack.cls` (클래스 `stack`)

' 문서별 스택:
' - docKey -> stack
Private gNavBackByDoc As Object ' Scripting.Dictionary (late-bound)
Private gNavForwardByDoc As Object ' Scripting.Dictionary (late-bound)

Private Const NAV_TMP_BM_PREFIX As String = "NavTemp_"

' 현재 위치에 임시 북마크 저장(Back stack push)
Public Sub PushBackwardBookmark(ByVal rng As Range)
    On Error GoTo SafeExit
    
    Dim doc As Document
    Set doc = rng.Document

    Dim s As stack
    Set s = GetBackStack(doc)
    
    Dim bmName As String
    bmName = CreateNavTmpBookmarkAtRange(rng)
    If bmName = "" Then GoTo SafeExit
    
    Call s.Push(bmName)
    
SafeExit:
End Sub

' 스택 pop 후 해당 임시 북마크로 이동. 성공하면 True
Public Function TryPopBackLocation() As Boolean
    On Error GoTo SafeExit

    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then GoTo SafeExit

    Dim backStack As stack
    Set backStack = GetBackStack(doc)
    If backStack.IsEmpty() Then GoTo SafeExit
    
    Dim targetBm As String
    targetBm = PopExistingNavTmpBookmarkName(doc, backStack)
    If targetBm = "" Then GoTo SafeExit

    ' Back으로 이동하기 전, 현재 위치를 Forward stack에 저장
    Dim cur As Range
    On Error Resume Next
    Set cur = Selection.Range
    On Error GoTo SafeExit
    If Not cur Is Nothing Then
        Call PushForwardLocation(cur)
    End If

    Call doc.Activate
    Call Selection.GoTo(What:=wdGoToBookmark, Name:=targetBm)
    Call SafeDeleteBookmark(doc, targetBm)
    
    TryPopBackLocation = True
    Exit Function
    
SafeExit:
    TryPopBackLocation = False
End Function

' 스택 pop 후 해당 임시 북마크로 이동(Forward). 성공하면 True
Public Function TryPopForwardLocation() As Boolean
    On Error GoTo SafeExit

    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then GoTo SafeExit

    Dim fwdStack As stack
    Set fwdStack = GetForwardStack(doc)
    If fwdStack.IsEmpty() Then GoTo SafeExit
    
    Dim targetBm As String
    targetBm = PopExistingNavTmpBookmarkName(doc, fwdStack)
    If targetBm = "" Then GoTo SafeExit

    ' Forward로 이동하기 전, 현재 위치를 Back stack에 저장
    Dim cur As Range
    On Error Resume Next
    Set cur = Selection.Range
    On Error GoTo SafeExit
    If Not cur Is Nothing Then
        Call PushBackwardBookmark(cur)
    End If

    Call doc.Activate
    Call Selection.GoTo(What:=wdGoToBookmark, Name:=targetBm)
    Call SafeDeleteBookmark(doc, targetBm)

    TryPopForwardLocation = True
    Exit Function

SafeExit:
    TryPopForwardLocation = False
End Function

' Alt+→ 동작 중 "앞으로가기 대상"을 스택에 저장(Forward stack push)
Private Sub PushForwardLocation(ByVal rng As Range)
    On Error GoTo SafeExit

    Dim doc As Document
    Set doc = rng.Document

    Dim s As stack
    Set s = GetForwardStack(doc)

    Dim bmName As String
    bmName = CreateNavTmpBookmarkAtRange(rng)
    If bmName = "" Then GoTo SafeExit
    
    Call s.Push(bmName)

SafeExit:
End Sub

' ======================
' stack clean-up
' ======================

' 뒤로 스택 지우기(문서별)
Public Sub ClearNavBackStack(ByVal doc As Document)
    On Error GoTo SafeExit
    
    If doc Is Nothing Then Exit Sub
   
    Call EnsureNavDictionaries()
    
    Dim docKey As String
    docKey = GetDocumentKey(doc)

    If gNavBackByDoc.Exists(docKey) Then
        Call DeleteNavTmpBookmarksInStack(doc, gNavBackByDoc(docKey))
        Call gNavBackByDoc.Remove(docKey)
    End If

    If gNavForwardByDoc.Exists(docKey) Then
        Call DeleteNavTmpBookmarksInStack(doc, gNavForwardByDoc(docKey))
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
        Call DeleteNavTmpBookmarksInStack(doc, gNavForwardByDoc(docKey))
        Call gNavForwardByDoc.Remove(docKey)
    End If

SafeExit:
End Sub

' ======================
' temp bookmark helpers
' ======================

' 현재 위치 Range에 임시 북마크를 생성하고 이름을 반환합니다. 실패 시 "".
Private Function CreateNavTmpBookmarkAtRange(ByVal rng As Range) As String
    On Error GoTo SafeExit
    
    If rng Is Nothing Then GoTo SafeExit
    
    Dim doc As Document
    Set doc = rng.Document
    If doc Is Nothing Then GoTo SafeExit
    
    Dim mark As Range
    Set mark = rng.Duplicate
    
    Dim bmName As String
    bmName = NAV_TMP_BM_PREFIX & GetUnixTimeMs()
    
    Call doc.Bookmarks.Add(Name:=bmName, Range:=mark)
    
    CreateNavTmpBookmarkAtRange = bmName
    Exit Function
    
SafeExit:
    CreateNavTmpBookmarkAtRange = ""
End Function

' stack에서 "실제로 존재하는" 임시 북마크 이름을 pop 해서 반환합니다. 없으면 "".
Private Function PopExistingNavTmpBookmarkName( _
    ByVal doc As Document, _
    ByVal s As stack _
) As String
    On Error GoTo SafeExit
    
    If doc Is Nothing Then GoTo SafeExit
    If s Is Nothing Then GoTo SafeExit
    
    Do While Not s.IsEmpty()
        Dim v As Variant
        v = s.Pop()
        
        If IsNull(v) Then Exit Do
        If VarType(v) <> vbString Then GoTo ContinueLoop
        
        Dim bmName As String
        bmName = CStr(v)
        If bmName = "" Then GoTo ContinueLoop
        If Left$(bmName, Len(NAV_TMP_BM_PREFIX)) <> NAV_TMP_BM_PREFIX Then
            GoTo ContinueLoop
        End If
        
        If doc.Bookmarks.Exists(bmName) Then
            PopExistingNavTmpBookmarkName = bmName
            Exit Function
        End If
        
ContinueLoop:
    Loop
    
SafeExit:
    PopExistingNavTmpBookmarkName = ""
End Function

Private Sub SafeDeleteBookmark(ByVal doc As Document, ByVal bmName As String)
    On Error GoTo SafeExit
    
    If doc Is Nothing Then GoTo SafeExit
    If bmName = "" Then GoTo SafeExit
    If Not doc.Bookmarks.Exists(bmName) Then GoTo SafeExit
    
    Call doc.Bookmarks(bmName).Delete
    
SafeExit:
End Sub

Private Sub DeleteNavTmpBookmarksInStack(ByVal doc As Document, ByVal s As stack)
    On Error GoTo SafeExit
    
    If doc Is Nothing Then GoTo SafeExit
    If s Is Nothing Then GoTo SafeExit
    
    Dim arr As Variant
    arr = s.ToArray()
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If VarType(arr(i)) = vbString Then
            Dim bmName As String
            bmName = CStr(arr(i))
            
            If Left$(bmName, Len(NAV_TMP_BM_PREFIX)) = NAV_TMP_BM_PREFIX Then
                Call SafeDeleteBookmark(doc, bmName)
            End If
        End If
    Next i
    
SafeExit:
End Sub

' ======================
' per-document stacks
' ======================

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

Private Sub EnsureNavDictionaries()
    If gNavBackByDoc Is Nothing Then
        Set gNavBackByDoc = CreateObject("Scripting.Dictionary")
    End If

    If gNavForwardByDoc Is Nothing Then
        Set gNavForwardByDoc = CreateObject("Scripting.Dictionary")
    End If
End Sub
