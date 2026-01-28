Option Explicit

' 클래스 모듈: stack
' ----------------------
' 범용 LIFO 스택 자료구조입니다.
'
' Public 멤버:
' - Push(item): 아이템을 스택에 추가
' - Pop(): 최상단 아이템 제거 후 반환(비어 있으면 Null)
' - Peek(): 최상단 아이템 반환(제거하지 않음, 비어 있으면 Null)
' - Clear(): 모든 아이템 제거
' - Count: 현재 아이템 개수
' - IsEmpty(): 비어 있으면 True
' - ToArray(): 하단->상단 순서의 1-based 배열 반환
'
' 메모:
' - 아이템 타입은 Variant(값/객체 모두 가능)로 취급합니다.
' - 객체는 `Call s.Push(obj)` 형태로 넣습니다. Pop/Peek은
'   `v = s.Pop() : If Not IsNull(v) Then Set obj = v`
'   패턴으로 꺼내야 안전합니다.

Private mItems As Collection

Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Function IsEmpty() As Boolean
    IsEmpty = (mItems.Count = 0)
End Function

Public Sub Clear()
    Set mItems = New Collection
End Sub

Public Sub Push(ByVal item As Variant)
    Call mItems.Add(item)
End Sub

Public Function Peek() As Variant
    If mItems.Count = 0 Then
        Peek = Null
        Exit Function
    End If

    If IsObject(mItems(mItems.Count)) Then
        Set Peek = mItems(mItems.Count)
    Else
        Peek = mItems(mItems.Count)
    End If
End Function

Public Function Pop() As Variant
    If mItems.Count = 0 Then
        Pop = Null
        Exit Function
    End If

    If IsObject(mItems(mItems.Count)) Then
        Set Pop = mItems(mItems.Count)
    Else
        Pop = mItems(mItems.Count)
    End If
    Call mItems.Remove(mItems.Count)
End Function

Public Function ToArray() As Variant
    Dim n As Long
    n = mItems.Count

    If n = 0 Then
        ToArray = VBA.Array()
        Exit Function
    End If

    Dim arr() As Variant
    ReDim arr(1 To n)

    Dim i As Long
    For i = 1 To n
        If IsObject(mItems(i)) Then
            Set arr(i) = mItems(i)
        Else
            arr(i) = mItems(i)
        End If
    Next i

    ToArray = arr
End Function
