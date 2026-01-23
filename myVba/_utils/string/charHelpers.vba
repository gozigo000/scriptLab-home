' 공백 문자 확인
Public Function IsWhitespaceChar(ByVal ch As String) As Boolean
    IsWhitespaceChar = (ch = " " Or ch = vbTab Or ch = ChrW$(160))
End Function

' 구분 문자 확인
Public Function IsDelimiterChar(ByVal ch As String) As Boolean
    Const DELIMITER_CHARS As String = " ~!@#$%&()-=+[]{}\/|'""<>?`^*,.;:【】"
    IsDelimiterChar = (InStr(DELIMITER_CHARS, ch) > 0)
End Function

Public Function IsAsciiNumber(ByVal ch As String) As Boolean
    Const code As Long = AscW(ch)
    IsAsciiNumber = (code >= 48 And code <= 57) ' 0-9
End Function

Public Function IsAsciiUpperCase(ByVal ch As String) As Boolean
    Const code As Long = AscW(ch)
    IsAsciiUpperCase = (code >= 65 And code <= 90) ' A-Z
End Function

Public Function IsAsciiLowerCase(ByVal ch As String) As Boolean
    Const code As Long = AscW(ch)
    IsAsciiLowerCase = (code >= 97 And code <= 122) ' a-z
End Function

Public Function IsAsciiAlphabet(ByVal ch As String) As Boolean
    Const code As Long = AscW(ch)
    IsAsciiAlphabet = (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) ' A-Z or a-z
End Function

Public Function IsAsciiAlphaNum(ByVal ch As String) As Boolean
    Const code As Long = AscW(ch)
    IsAsciiAlphaNum = (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) ' 0-9 or A-Z or a-z
End Function