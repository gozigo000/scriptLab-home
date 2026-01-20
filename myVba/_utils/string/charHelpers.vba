' (MARK) 공백 문자 판정
Public Function IsWhitespaceChar(ByVal ch As String) As Boolean
    If ch = " " Or ch = vbTab Or ch = ChrW$(160) Then
        IsWhitespaceChar = True
    Else
        IsWhitespaceChar = False
    End If
End Function