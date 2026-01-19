Option Explicit

' ======================
' Regex Pattern Test Helper
' ======================
'
' - 문자열이 정규식 패턴에 매칭되는지 테스트
'
' 사용 방법:
' - TestRegexPattern(문자열, 정규식 패턴)
'   예) TestRegexPattern("hello", "hello") -> True
'   예) TestRegexPattern("hello", "world") -> False
'   예) TestRegexPattern("", "hello") -> False
'   예) TestRegexPattern("hello", "") -> False
'
' 예외 처리:
' - 문자열이 비어있거나 패턴이 비어있으면 False 반환
' - 문자열이 정규식 패턴에 매칭되면 True 반환
' - 문자열이 정규식 패턴에 매칭되지 않으면 False 반환
'
Public Function TestRegex(ByVal str As String, ByVal pattern As String) As Boolean
    On Error GoTo ReturnFalse
    
    str = Trim$(str)
    If Len(str) = 0 Then GoTo ReturnFalse
    
    ' VBScript.RegExp (late binding) - 별도 참조 설정 없이 사용 가능
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = pattern
    
    TestRegex = CBool(re.Test(str))
    Exit Function
    
ReturnFalse:
    TestRegex = False
End Function
