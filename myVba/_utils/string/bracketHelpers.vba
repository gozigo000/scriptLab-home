' 문자가 괄호인지 확인하는 함수
Public Function IsBracket(char As String) As Boolean
    IsBracket = (char = "(" Or char = ")" Or _
                 char = "[" Or char = "]" Or _
                 char = "{" Or char = "}")
End Function

' 여는 괄호인지 확인하는 함수
Public Function IsOpenBracket(char As String) As Boolean
    IsOpenBracket = (char = "(" Or char = "[" Or char = "{")
End Function

' 닫는 괄호인지 확인하는 함수
Public Function IsCloseBracket(char As String) As Boolean
    IsCloseBracket = (char = ")" Or char = "]" Or char = "}")
End Function


' 괄호의 종류를 반환하는 함수 ("()", "[]", "{}")
Public Function GetBracketType(char As String) As String
    If char = "(" Or char = ")" Then
        GetBracketType = "()"
    ElseIf char = "[" Or char = "]" Then
        GetBracketType = "[]"
    ElseIf char = "{" Or char = "}" Then
        GetBracketType = "{}"
    Else
        GetBracketType = ""
    End If
End Function