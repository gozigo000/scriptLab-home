Public Function GetUnixTimeS() As String
    ' Unix time: 1970-01-01 00:00:00 기준 경과 초
    On Error GoTo SafeExit
    
#If VBA7 Then
    Dim seconds As LongLong
    seconds = DateDiff("s", #1/1/1970#, Now)
    GetUnixTimeS = CStr(seconds)
    Exit Function
#Else
    Dim seconds As Long
    seconds = DateDiff("s", #1/1/1970#, Now)
    GetUnixTimeS = CStr(seconds)
    Exit Function
#End If
    
SafeExit:
    GetUnixTimeS = ""
End Function

Public Function GetUnixTimeMs() As String
    ' Unix time: 1970-01-01 00:00:00 기준 경과 밀리초
    ' - VBA의 Now는 초 해상도인 경우가 많아, ms는 Timer의 소수부를 사용합니다.
    ' - 같은 ms 내 충돌 방지는 호출부에서 seq를 붙이도록 합니다.
    On Error GoTo SafeExit
    
    Dim ms As Long
    ms = CLng((Timer - Fix(Timer)) * 1000)
    If ms < 0 Then ms = 0
    If ms > 999 Then ms = 999
    
#If VBA7 Then
    Dim seconds As LongLong
    seconds = DateDiff("s", #1/1/1970#, Now)
    GetUnixTimeMs = CStr(seconds * 1000 + ms)
    Exit Function
#Else
    Dim seconds As Long
    seconds = DateDiff("s", #1/1/1970#, Now)
    
    ' Long 범위를 넘길 수 있어 문자열 결합으로 반환합니다.
    GetUnixTimeMs = CStr(seconds) & Right$("000" & CStr(ms), 3)
    Exit Function
#End If
    
SafeExit:
    GetUnixTimeMs = ""
End Function