Option Explicit

' ======================
' Variable Case Helpers
' ======================
'
' - camelCase:  첫 글자 소문자, 언더바 없음, 이후 단어 경계는 대문자 (숫자 허용)
'   예) myVar, my2Var, myURL
'
' - PascalCase: 첫 글자 대문자, 언더바 없음, 이후 단어 경계는 대문자 (숫자 허용)
'   예) MyVar, My2Var, URLValue
'
' - snake_case: 전부 소문자/숫자 + 단어 구분은 '_' (연속/양끝 '_' 금지)
'   예) my_var, my2_var3, my_url
'
' 주의:
' - 빈 문자열/공백 포함/특수문자 포함은 False
' - 선행/후행 '_'(예: _var, var_)은 snake_case로 보지 않음

Public Function IsCamelCase(ByVal str As String) As Boolean
    IsCamelCase = TestRegex(str, "^[a-z][a-z0-9]*([A-Z][a-z0-9]*)+$")
End Function

Public Function IsPascalCase(ByVal str As String) As Boolean
    IsPascalCase = TestRegex(str, "^[A-Z][a-z0-9]+([A-Z][a-z0-9]*)+$")
End Function

Public Function IsSnakeCase(ByVal str As String) As Boolean
    IsSnakeCase = TestRegex(str, "^[a-z][a-z0-9]*(_[a-z0-9]+)+$")
End Function

Public Function IsScreamingSnakeCase(ByVal str As String) As Boolean
    IsScreamingSnakeCase = TestRegex(str, "^[A-Z][A-Z0-9]*(_[A-Z0-9]+)+$")
End Function

