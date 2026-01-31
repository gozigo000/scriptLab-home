Option Explicit

' ============================================================
' 모듈: outlineCache2
' 역할: Range.GoTo 기반 "현재 제목(개요)" 탐색 유틸
' Public 함수/서브:
'   - GetCurrentHeadingTitle2: 현재 위치의 제목 문자열 반환
'   - GetCurrentHeadingLevel2: 현재 위치의 제목 레벨 반환
'   - GetCurrentHeadingRange2: 현재 위치의 제목 구간(Range) 반환
'   - (디버깅용) a_ShowOutlineHeadingInfo2: 현재 위치의 제목 정보 알림
' ============================================================
'
' 사용 예:
' - title = GetCurrentHeadingTitle2(Selection.Range)
'
' 참고:
' - https://learn.microsoft.com/ko-kr/office/vba/api/word.range.goto
'

' 현재 커서가 속한 문단에서 위쪽으로 가장 가까운
' "제목(탐색창에 표시되는 개요 수준)"을 찾아 반환
' - rng: 현재 커서가 속한 문단의 Range
' - 반환값: "제목" (또는 "1.2.3 제목" 처럼 목록번호가 있으면 포함)
Public Function GetCurrentHeadingTitle2( _
    ByVal rng As Range _
) As String
    On Error GoTo SafeExit

    If rng Is Nothing Then GoTo SafeExit

    Dim p As Paragraph
    Set p = GetCurrentHeadingParagraph(rng)
    If p Is Nothing Then GoTo SafeExit

    Dim title As String
    title = NormalizeText(p.Range.Text)

    Dim num As String
    num = GetParagraphListNumber(p)

    If num <> "" Then
        title = num & " " & title
    End If

    GetCurrentHeadingTitle2 = title
    Exit Function

SafeExit:
    GetCurrentHeadingTitle2 = ""
End Function

' 현재 커서가 속한 문단의 "제목 레벨(OutlineLevel)"을 반환
' - rng: 현재 커서가 속한 문단의 Range
' - 반환값: WdOutlineLevel 값을 Long으로 반환(없으면 0)
Public Function GetCurrentHeadingLevel2( _
    ByVal rng As Range _
) As Long
    On Error GoTo SafeExit

    If rng Is Nothing Then GoTo SafeExit

    Dim p As Paragraph
    Set p = GetCurrentHeadingParagraph(rng)
    If p Is Nothing Then GoTo SafeExit

    GetCurrentHeadingLevel2 = CLng(p.OutlineLevel)
    Exit Function

SafeExit:
    GetCurrentHeadingLevel2 = 0
End Function

' 현재 위치 기준:
' - 위로 가장 가까운 제목(헤딩) 문단 시작 ~
' - 아래로 가장 가까운 다음 제목(헤딩) 문단 시작 직전
' 을 Range로 반환합니다. (End는 다음 제목 Start로 설정)
' - 제목을 찾지 못하면 Nothing 반환
Public Function GetCurrentHeadingRange2( _
    ByVal rng As Range _
) As Range
    On Error GoTo SafeExit

    If rng Is Nothing Then GoTo SafeExit

    Dim p As Paragraph
    Set p = GetCurrentHeadingParagraph(rng)
    If p Is Nothing Then GoTo SafeExit

    Dim doc As Document
    Set doc = rng.Document

    Dim headingStart As Long
    headingStart = p.Range.Start

    Dim nextHeadingStart As Long
    nextHeadingStart = GetNextHeadingStart(p)

    Dim endExclusive As Long
    endExclusive = doc.Content.End
    If nextHeadingStart > 0 And nextHeadingStart <= doc.Content.End Then
        If nextHeadingStart > headingStart Then endExclusive = nextHeadingStart
    End If

    If endExclusive < headingStart Then endExclusive = headingStart
    Set GetCurrentHeadingRange2 = doc.Range(headingStart, endExclusive)
    Exit Function

SafeExit:
    Set GetCurrentHeadingRange2 = Nothing
End Function

' (MARK) (디버깅용) 현재 위치의 "제목(개요)" 정보 알림
Public Sub a_ShowOutlineHeadingInfo2()
    On Error GoTo ErrorHandler

    Dim selRng As Range
    Set selRng = Selection.Range
    If selRng Is Nothing Then
        VBA.MsgBox "Selection.Range를 가져올 수 없습니다.", vbInformation, "현재 제목 정보"
        Exit Sub
    End If

    Dim headingTitle As String
    headingTitle = GetCurrentHeadingTitle2(selRng)

    Dim headingLevel As Long
    headingLevel = GetCurrentHeadingLevel2(selRng)

    Dim headingRng As Range
    Set headingRng = GetCurrentHeadingRange2(selRng)

    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim rangeLen As Long

    If headingRng Is Nothing Then
        rangeStart = 0
        rangeEnd = 0
        rangeLen = 0
    Else
        rangeStart = headingRng.Start
        rangeEnd = headingRng.End
        rangeLen = rangeEnd - rangeStart
    End If

    Dim msg As String
    msg = ""
    msg = msg & "제목: " & IIf(headingTitle = "", "(없음)", headingTitle) & vbCrLf
    msg = msg & "레벨: " & CStr(headingLevel) & vbCrLf
    msg = msg & "구간: [" & CStr(rangeStart) & ", " & CStr(rangeEnd) & ") " & _
        "(len=" & CStr(rangeLen) & ")" & vbCrLf

    VBA.MsgBox msg, vbInformation, "현재 제목 정보"
    Exit Sub

ErrorHandler:
    VBA.MsgBox "오류: " & Err.Description, vbCritical, "현재 제목 정보"
End Sub


' ===========================
' 내부 구현: Range.GoTo 기반
' ===========================

' rng가 속한 "현재 제목 문단"을 반환합니다.
Private Function GetCurrentHeadingParagraph( _
    ByVal rng As Range _
) As Paragraph
    On Error GoTo SafeExit

    If rng Is Nothing Then GoTo SafeExit
    If rng.Paragraphs.Count <= 0 Then GoTo SafeExit

    Dim p As Paragraph
    Set p = rng.Paragraphs(1)

    ' 현재 문단이 제목이면 그대로 반환
    If p.OutlineLevel <> wdOutlineLevelBodyText Then
        Set GetCurrentHeadingParagraph = p
        Exit Function
    End If

    ' Range.GoTo(wdGoToHeading, wdGoToPrevious)로 이전 제목을 찾음
    Dim probe As Range
    Set probe = rng.Duplicate
    probe.Collapse wdCollapseStart

    Dim headRng As Range
    On Error Resume Next
    Set headRng = probe.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
    On Error GoTo SafeExit

    If headRng Is Nothing Then GoTo SafeExit

    Set p = headRng.Paragraphs(1)
    If p.OutlineLevel = wdOutlineLevelBodyText Then GoTo SafeExit

    Set GetCurrentHeadingParagraph = p
    Exit Function

SafeExit:
    Set GetCurrentHeadingParagraph = Nothing
End Function

' headingPara 다음에 나오는 "다음 제목" 시작점을 반환합니다.
' - 없으면 0 반환
Private Function GetNextHeadingStart( _
    ByVal headingPara As Paragraph _
) As Long
    On Error GoTo SafeExit

    If headingPara Is Nothing Then GoTo SafeExit

    Dim seekRng As Range
    Set seekRng = headingPara.Range.Duplicate
    seekRng.Collapse(wdCollapseEnd)

    Dim nextRng As Range
    On Error Resume Next
    Set nextRng = seekRng.GoTo(What:=wdGoToHeading, Which:=wdGoToNext)
    On Error GoTo SafeExit

    If nextRng Is Nothing Then GoTo SafeExit

    Dim nextStart As Long
    nextStart = nextRng.Start

    If nextStart <= headingPara.Range.Start Then GoTo SafeExit

    GetNextHeadingStart = nextStart
    Exit Function

SafeExit:
    GetNextHeadingStart = 0
End Function

' 문단 텍스트 정규화 (한 줄로)
Private Function NormalizeText(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW$(160), " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeText = Trim$(s)
End Function

' 문단에 표시되는 "목록 번호(예: 1.2.3)" 문자열 반환
' - 다단계 번호/헤딩 번호 포함
' - 번호가 없거나 접근 실패 시 "" 반환
Private Function GetParagraphListNumber(ByVal p As Paragraph) As String
    On Error Resume Next

    Dim s As String
    s = p.Range.ListFormat.ListString

    If Err.Number <> 0 Then
        Err.Clear
        s = ""
    End If

    On Error GoTo 0
    GetParagraphListNumber = Trim$(s)
End Function

