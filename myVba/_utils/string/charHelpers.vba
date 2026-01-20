' (MARK) 공백 문자 판정
Public Function IsWhitespaceChar(ByVal ch As String) As Boolean
    If ch = " " Or ch = vbTab Or ch = ChrW$(160) Then
        IsWhitespaceChar = True
    Else
        IsWhitespaceChar = False
    End If
End Function

Public Function IsLowerAsciiLetter(ByVal ch As String) As Boolean
    Const code As Long = AscW(ch)
    IsLowerAsciiLetter = (code >= 97 And code <= 122) ' a-z
End Function

Public Function IsUpperAsciiLetter(ByVal ch As String) As Boolean
    Const code As Long = AscW(ch)
    IsUpperAsciiLetter = (code >= 65 And code <= 90) ' A-Z
End Function

Public Function IsAsciiDigit(ByVal ch As String) As Boolean
    Const code As Long = AscW(ch)
    IsAsciiDigit = (code >= 48 And code <= 57) ' 0-9
End Function