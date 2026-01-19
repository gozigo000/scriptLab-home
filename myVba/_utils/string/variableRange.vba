Option Explicit

' ======================
' Variable Range Helpers
' ======================
'
' - "코딩에서 일반적으로 사용하는 변수명"을 단어로 취급: [A-Za-z0-9_]
' - 커서(삽입점) 위치를 입력으로 받아, 해당 변수 범위(Range)를 반환
' - 변수 위/중간/끝에 커서가 있어도 전체 변수 Range를 반환
' - 변수가 아니면 Nothing 반환

' (MARK) 커서 Range(삽입점) 기준 변수 Range 반환
Public Function GetVariableRangeAtInsPnt(ByVal doc As Document, ByVal cursorRng As Range) As Range
    On Error GoTo ReturnNothing
    If doc Is Nothing Then GoTo ReturnNothing
    If cursorRng Is Nothing Then GoTo ReturnNothing
    
    If cursorRng.Type <> wdSelectionIP Then GoTo ReturnNothing
    
    Dim pos As Long
    pos = cursorRng.Start
    
    Set GetVariableRangeAtInsPnt = GetVariableRangeAtPos(doc, pos)
    Exit Function
    
ReturnNothing:
    Set GetVariableRangeAtInsPnt = Nothing
End Function

' (MARK) (doc, pos) 기준 변수 Range 반환
' - pos는 "커서 위치" (Range.Start)처럼 취급
Public Function GetVariableRangeAtPos(ByVal doc As Document, ByVal pos As Long) As Range
    On Error GoTo SafeExit
    If doc Is Nothing Then GoTo SafeExit
    
    Dim docEnd As Long
    docEnd = doc.Content.End
    
    If pos < 0 Then pos = 0
    If pos > docEnd Then pos = docEnd
    
    ' 1) 기준 문자 위치 결정
    ' - pos가 변수 문자 위면 pos
    ' - 아니면 pos-1이 변수 문자면 pos-1 (단어 끝에 커서가 있는 케이스)
    Dim basePos As Long
    basePos = -1
    
    If pos < docEnd Then
        If IsVariableChar(GetCharAtPos(doc, pos)) Then
            basePos = pos
        End If
    End If
    
    If basePos = -1 Then
        If pos > 0 Then
            If IsVariableChar(GetCharAtPos(doc, pos - 1)) Then
                basePos = pos - 1
            End If
        End If
    End If
    
    If basePos = -1 Then GoTo SafeExit
    
    ' 2) 좌/우로 확장해서 변수 범위 계산
    Dim startPos As Long
    Dim endPos As Long
    startPos = basePos
    endPos = basePos + 1 ' end는 exclusive
    
    ' 왼쪽 확장: startPos-1 이 변수 문자면 계속 확장
    Do While startPos > 0
        If Not IsVariableChar(GetCharAtPos(doc, startPos - 1)) Then Exit Do
        startPos = startPos - 1
    Loop
    
    ' 오른쪽 확장: endPos 가 변수 문자면 계속 확장
    Do While endPos < docEnd
        If Not IsVariableChar(GetCharAtPos(doc, endPos)) Then Exit Do
        endPos = endPos + 1
    Loop
    
    If endPos <= startPos Then GoTo SafeExit
    
    Set GetVariableRangeAtPos = doc.Range(startPos, endPos)
    Exit Function
    
SafeExit:
    Set GetVariableRangeAtPos = Nothing
End Function

' ======================
' Selection / Text helpers
' ======================
'
' - 선택 영역(Selection) 또는 임의 Range의 텍스트를 문자열로 반환
' - 선택이 없거나(커서만 있는 상태) 입력이 Nothing이면 "" 반환

' (MARK) 선택(Selection) 범위의 텍스트 반환
Public Function GetSelectionText(Optional ByVal sel As Selection) As String
    On Error GoTo ReturnEmptyString
    
    Dim s As Selection
    If sel Is Nothing Then
        Set s = Selection
    Else
        Set s = sel
    End If
    
    If s Is Nothing Then GoTo ReturnEmptyString
    If s.Range Is Nothing Then GoTo ReturnEmptyString
    
    ' 커서(IP)만 있는 상태면 "선택범위"가 없으므로 빈 문자열
    If s.Type = wdSelectionIP Then GoTo ReturnEmptyString
    If s.Range.Start = s.Range.End Then GoTo ReturnEmptyString
    
    GetSelectionText = CStr(s.Range.Text)
    Exit Function
    
ReturnEmptyString:
    GetSelectionText = ""
End Function

' (MARK) Range의 텍스트 반환 (Nothing이면 "")
Public Function GetRangeText(ByVal rng As Range) As String
    On Error GoTo SafeExit
    If rng Is Nothing Then GoTo SafeExit
    
    GetRangeText = CStr(rng.Text)
    Exit Function
    
SafeExit:
    GetRangeText = ""
End Function

' (MARK) 위치(pos)의 변수명(String) 반환
' - 변수가 아니면 "" 반환
Public Function GetVariableStringAtPos(ByVal doc As Document, ByVal pos As Long) As String
    On Error GoTo SafeExit
    
    Dim varRng As Range
    Set varRng = GetVariableRangeAtPos(doc, pos)
    If varRng Is Nothing Then GoTo SafeExit
    
    GetVariableStringAtPos = CStr(varRng.Text)
    Exit Function
    
SafeExit:
    GetVariableStringAtPos = ""
End Function

' ----------------------
' Private helpers
' ----------------------

Private Function GetCharAtPos(ByVal doc As Document, ByVal pos As Long) As String
    On Error GoTo SafeExit
    If doc Is Nothing Then GoTo SafeExit
    
    Dim docEnd As Long
    docEnd = doc.Content.End
    If pos < 0 Or pos >= docEnd Then GoTo SafeExit
    
    Dim r As Range
    Set r = doc.Range(pos, pos + 1)
    GetCharAtPos = r.Text
    Exit Function
    
SafeExit:
    GetCharAtPos = ""
End Function

Private Function IsVariableChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsVariableChar = False
        Exit Function
    End If
    
    ' VBA Like: 단일 문자 패턴 매칭
    IsVariableChar = (ch Like "[A-Za-z0-9_]")
End Function
