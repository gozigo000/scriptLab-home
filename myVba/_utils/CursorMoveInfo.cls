Option Explicit

' (MARK) 커서 이동 정보 객체
' ----------------------
' 커서 이동 시점의 스냅샷을 담는 용도입니다.
' - VBA 편집기에서 "삽입 > 클래스 모듈" 추가 후
'   Name을 "CursorMoveInfo"로 설정하고 이 내용을 붙여넣으세요.

Private mPageNo As Long
Private mWordText As String
Private mBookmarkNames As String
Private mSubAddress As String
Private mPosition As Long

Public Property Get PageNo() As Long
    PageNo = mPageNo
End Property

Public Property Let PageNo(ByVal v As Long)
    mPageNo = v
End Property

Public Property Get WordText() As String
    WordText = mWordText
End Property

Public Property Let WordText(ByVal v As String)
    mWordText = v
End Property

Public Property Get BookmarkNames() As String
    BookmarkNames = mBookmarkNames
End Property

Public Property Let BookmarkNames(ByVal v As String)
    mBookmarkNames = v
End Property

Public Property Get SubAddress() As String
    SubAddress = mSubAddress
End Property

Public Property Let SubAddress(ByVal v As String)
    mSubAddress = v
End Property

Public Property Get Position() As Long
    Position = mPosition
End Property

Public Property Let Position(ByVal v As Long)
    mPosition = v
End Property

