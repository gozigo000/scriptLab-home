Option Explicit

' (MARK) 선택 영역 자동 검색 섹션
' ----------------------
' 페이지 로드 시 이벤트 핸들러 등록 (기본적으로 비활성화 상태)
' 사용자가 버튼을 클릭하면 활성화됨
'
' 사용 방법:
' 1. 이 모듈을 Word VBA 프로젝트에 추가합니다.
' 2. MsgBox 모듈을 Word VBA 프로젝트에 추가합니다.
'    (MsgBox.vba 파일 참조)
' 3. clsAppEvents 클래스 모듈을 Word VBA 프로젝트에 추가합니다.
'    (clsAppEvents.cls 파일 참조)
' 4. ThisDocument 모듈에 다음 코드를 추가합니다:
'
'    Dim myAppEvents As clsAppEvents
'
'    Private Sub Document_Open()
'        Call InitializeCurrWordHighlighter
'        Set myAppEvents = New clsAppEvents
'        Set myAppEvents.appWord = Word.Application
'    End Sub
'
' 5. 기능을 활성화하려면 ToggleCurrWordHighlighter 서브루틴을 호출합니다.
'    예: 매크로 버튼이나 단축키에 ToggleCurrWordHighlighter 할당

' 이전 하이라이트 정보를 담는 타입
Public Type TPrevHighlightInfo
    Word As String
    ScopeStart As Long ' 직전에 하이라이트를 적용한 "섹션" 범위 시작
    ScopeEnd As Long   ' 직전에 하이라이트를 적용한 "섹션" 범위 끝(End, exclusive)
End Type

' 모듈 레벨 변수
Public isCurrWordHighlighterEnabled As Boolean ' 기능 ON/OFF 토글 (True=활성화)
Public isProcessingSelectionChange As Boolean ' 무한루프 방지 플래그
Public gPrevHighlight As TPrevHighlightInfo

' (MARK) 하이라이트 스타일
' 표 내부에서 Shading은 레이아웃 재계산을 자주 유발하므로,
' 텍스트 하이라이트(HighlightColorIndex)를 사용합니다.
Private Const CURRWORD_HIGHLIGHT_COLOR As Long = wdBrightGreen

' (MARK) 초기화
Public Sub InitializeCurrWordHighlighter()
    isCurrWordHighlighterEnabled = False ' 초기 상태: 비활성화
    isProcessingSelectionChange = False
    gPrevHighlight.Word = ""
    gPrevHighlight.ScopeStart = 0
    gPrevHighlight.ScopeEnd = 0
End Sub

' (MARK) 단어 하이라이트 기능 토글
Public Sub ToggleCurrWordHighlighter()
    isCurrWordHighlighterEnabled = Not isCurrWordHighlighterEnabled
    
    If isCurrWordHighlighterEnabled Then
        ' 기능 활성화
        Call showMsg("단어 하이라이트 기능이 활성화되었습니다.", "알림", vbInformation, 500)
    Else
        ' 기능 비활성화
        ' 이전 하이라이트 제거
        If gPrevHighlight.Word <> "" Then
            Call RemoveHighlight(ActiveDocument, gPrevHighlight.Word, GetPreviousHighlightedScopeRange(ActiveDocument))
            gPrevHighlight.Word = ""
            gPrevHighlight.ScopeStart = 0
            gPrevHighlight.ScopeEnd = 0
        End If
        Call showMsg("단어 하이라이트 기능이 비활성화되었습니다.", "알림", vbInformation, 500)
    End If
End Sub

' (MARK) 단어 하이라이트
Public Sub HighlightCurrWord(ByVal targetRange As Range)
    If targetRange Is Nothing Then Exit Sub
    ' 기능이 비활성화되어 있으면 종료
    If Not isCurrWordHighlighterEnabled Then Exit Sub
    ' 텍스트가 선택(드래그)된 경우에는 하이라이트하지 않음
    If targetRange.Start <> targetRange.End Then Exit Sub
    ' 무한루프 방지: 이미 처리 중이면 종료
    If isProcessingSelectionChange Then Exit Sub
    
    ' 처리 중 플래그 설정
    isProcessingSelectionChange = True
    
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Set doc = targetRange.Document
    
    ' 현재 섹션(헤딩 단위) 범위로 검색/하이라이트 범위를 제한
    Dim scopeRange As Range
    Set scopeRange = GetCurrentHeadingRange(targetRange)
    If scopeRange Is Nothing Then
        ' 헤딩이 없는 문서/구간이면 기존처럼 문서 전체를 하나의 섹션으로 간주
        Set scopeRange = doc.Content
    End If
    
    Dim varRng As Range
    Set varRng = GetVariableRangeAtPos(doc, targetRange.Start)
    
    Dim currentWord As String
    currentWord = GetRangeText(varRng)

    ' 이전 단어와 동일하면 유지(단, "같은 섹션"일 때만)
    If currentWord = gPrevHighlight.Word Then
        If gPrevHighlight.ScopeStart = scopeRange.Start And _
            gPrevHighlight.ScopeEnd = scopeRange.End _
        Then
            GoTo Cleanup
        End If
        ' 단어는 같지만 섹션이 바뀐 경우:
        ' - 이전 섹션의 하이라이트는 지우고
        ' - 현재 섹션에 다시 적용해야 함 (아래 로직으로 진행)
    End If

    ' variableCase.vba 유틸: 케이스 판별 (camel/snake/pascal)
    If currentWord = "" Or Not ( _
        IsCamelCase(currentWord) Or _
        IsPascalCase(currentWord) Or _
        IsSnakeCase(currentWord) Or _
        IsScreamingSnakeCase(currentWord) _
    ) Then
        If gPrevHighlight.Word <> "" Then
            Call BeginCustomUndoRecord()
            Call RemoveHighlight(doc, gPrevHighlight.Word, GetPreviousHighlightedScopeRange(doc))
            gPrevHighlight.Word = ""
            gPrevHighlight.ScopeStart = 0
            gPrevHighlight.ScopeEnd = 0
        End If
        GoTo Cleanup
    End If
    
    
    Dim findRange As Range
    
    ' 이전 하이라이트 제거
    If gPrevHighlight.Word <> "" Then
        Call BeginCustomUndoRecord()
        Call RemoveHighlight(doc, gPrevHighlight.Word, GetPreviousHighlightedScopeRange(doc))
    End If
    
    Application.ScreenUpdating = False
    
    ' 현재 섹션에 하이라이트 적용도 동일 UndoRecord로 포함
    Call BeginCustomUndoRecord()
    
    ' 2-PASS
    ' 1) 전체 검색 범위를 순회하며 매칭 영역(start/end) 수집
    Dim spans As Collection
    Set spans = CollectMatchSpans(scopeRange, currentWord)
    
    ' 2) 수집된 영역에 대해 하이라이트 적용
    Dim span As Variant
    For Each span In spans
        Dim s As Long
        Dim e As Long
        s = CLng(span(0))
        e = CLng(span(1))
        
        If e > s Then
            Dim hit As Range
            Set hit = doc.Range(s, e)
            hit.HighlightColorIndex = CURRWORD_HIGHLIGHT_COLOR
        End If
    Next span
    
    gPrevHighlight.Word = currentWord
    gPrevHighlight.ScopeStart = scopeRange.Start
    gPrevHighlight.ScopeEnd = scopeRange.End
    
Cleanup:
    On Error Resume Next
    Call EndCustomUndoRecord()
    Application.ScreenUpdating = True
    isProcessingSelectionChange = False
    Exit Sub
    
ErrorHandler:
    Debug.Print "선택 영역 검색 및 하이라이트 적용 중 오류: " & Err.Description
    Resume Cleanup
End Sub

' (MARK) 이전 하이라이트 제거
Public Sub RemoveHighlight( _
    ByVal doc As Document, _
    ByVal searchText As String, _
    Optional ByVal scopeRange As Range _
)
    On Error GoTo ErrorHandler
    
    If doc Is Nothing Then Exit Sub
    If searchText = "" Then Exit Sub
    
    Dim findRange As Range
    
    Application.ScreenUpdating = False
    
    ' 검색/제거 범위: 기본은 문서 전체(호출자가 범위를 주면 그 범위만)
    If scopeRange Is Nothing Then
        Set scopeRange = doc.Content
    End If
    
    Dim spans As Collection
    Set spans = CollectMatchSpans(scopeRange, searchText)
    
    Dim span As Variant
    For Each span In spans
        Dim s As Long
        Dim e As Long
        s = CLng(span(0))
        e = CLng(span(1))
        
        If e > s Then
            Dim hit As Range
            Set hit = doc.Range(s, e)
            hit.HighlightColorIndex = wdNoHighlight
        End If
    Next span
    
Cleanup:
    On Error Resume Next
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Debug.Print "하이라이트 제거 중 오류: " & Err.Description
    Resume Cleanup
End Sub

' (MARK) 직전에 하이라이트한 섹션 범위를 복원 (없으면 doc.Content)
Private Function GetPreviousHighlightedScopeRange(ByVal doc As Document) As Range
    On Error GoTo Fallback
    
    If doc Is Nothing Then GoTo Fallback
    If gPrevHighlight.ScopeStart <= 0 Then GoTo Fallback
    If gPrevHighlight.ScopeEnd <= 0 Then GoTo Fallback
    If gPrevHighlight.ScopeEnd < gPrevHighlight.ScopeStart Then GoTo Fallback
    
    Dim s As Long
    Dim e As Long
    s = gPrevHighlight.ScopeStart
    e = gPrevHighlight.ScopeEnd
    
    ' 문서 편집 등으로 값이 어긋났을 수 있어 클램프
    If s < doc.Content.Start Then s = doc.Content.Start
    If e > doc.Content.End Then e = doc.Content.End
    If e < s Then GoTo Fallback
    
    Set GetPreviousHighlightedScopeRange = doc.Range(s, e)
    Exit Function
    
Fallback:
    If doc Is Nothing Then
        Set GetPreviousHighlightedScopeRange = Nothing
    Else
        Set GetPreviousHighlightedScopeRange = doc.Content
    End If
End Function

' ======================
' 커서 단어 하이라이트용 유틸
' ======================

' Find로 잡힌 구간이 "부분 문자열"이 아닌지(좌/우 경계가 식별자 문자가 아닌지) 확인
Private Function IsBoundaryMatch(ByVal foundRange As Range, ByVal scopeRange As Range) As Boolean
    On Error GoTo SafeExit
    
    Dim beforeCh As String
    Dim afterCh As String
    Dim doc As Document
    
    Set doc = scopeRange.Document
    
    ' 왼쪽 문자
    If foundRange.Start > scopeRange.Start Then
        beforeCh = doc.Range(foundRange.Start - 1, foundRange.Start).Text
        If IsVariableChar(beforeCh) Then GoTo SafeExit
    End If
    
    ' 오른쪽 문자
    If foundRange.End < scopeRange.End Then
        afterCh = doc.Range(foundRange.End, foundRange.End + 1).Text
        If IsVariableChar(afterCh) Then GoTo SafeExit
    End If
    
    IsBoundaryMatch = True
    Exit Function
    
SafeExit:
    IsBoundaryMatch = False
End Function

' ======================
' 2-pass Find helpers
' ======================

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then
        MaxLong = a
    Else
        MaxLong = b
    End If
End Function

Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then
        MinLong = a
    Else
        MinLong = b
    End If
End Function

' (MARK) 표 셀 전용 수집: Find 대신 InStr 사용
' - Word Range.Find는 표/셀 컨텍스트에서 "stuck(같은 매치 반복)"이 발생할 수 있어
'   표 내부에서는 문자열 검색(InStr)로 위치를 계산합니다.
Private Sub CollectInCellByInStr( _
    ByVal cellSeg As Range, _
    ByVal scopeRange As Range, _
    ByVal searchText As String, _
    ByVal seen As Object, _
    ByVal results As Collection _
)
    On Error GoTo SafeExit
    
    If cellSeg Is Nothing Then Exit Sub
    If scopeRange Is Nothing Then Exit Sub
    If searchText = "" Then Exit Sub
    
    Dim txt As String
    txt = CStr(cellSeg.Text)
    If txt = "" Then Exit Sub
    
    Dim n As Long
    n = Len(searchText)
    If n <= 0 Then Exit Sub
    
    Dim i As Long
    i = 1
    
    Do While i > 0
        i = InStr(i, txt, searchText, vbBinaryCompare)
        If i <= 0 Then Exit Do
        
        Dim absS As Long
        Dim absE As Long
        absS = cellSeg.Start + (i - 1)
        absE = absS + n
        
        If absE > absS Then
            Dim hit As Range
            Set hit = cellSeg.Document.Range(absS, absE)
            
            If IsBoundaryMatch(hit, scopeRange) Then
                Dim k As String
                k = CStr(absS) & ":" & CStr(absE)
                If Not seen.Exists(k) Then
                    seen.Add k, True
                    results.Add Array(absS, absE)
                End If
            End If
        End If
        
        ' 다음 후보로 전진 (겹침 매치는 필요 없으므로 길이만큼 점프)
        i = i + n
        If i > Len(txt) Then Exit Do
    Loop
    
SafeExit:
End Sub

' scopeRange 전체를 "순회"하면서(표는 셀 단위, 표 밖은 문단 단위)
' searchText 매칭 영역(start/end)을 모아 반환합니다.
'
' 반환 형태:
' - Collection 각 원소는 Variant(Array(start As Long, end As Long))
Private Function CollectMatchSpans(ByVal scopeRange As Range, ByVal searchText As String) As Collection
    On Error GoTo SafeExit
    
    Dim results As Collection
    Set results = New Collection
    
    If scopeRange Is Nothing Then
        Set CollectMatchSpans = results
        Exit Function
    End If
    If searchText = "" Then
        Set CollectMatchSpans = results
        Exit Function
    End If
    
    Dim doc As Document
    Set doc = scopeRange.Document
    
    ' 중복 제거용
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    
    Dim scopeS As Long
    Dim scopeE As Long
    scopeS = scopeRange.Start
    scopeE = scopeRange.End
    
    Dim tableCount As Long
    On Error Resume Next
    tableCount = scopeRange.Tables.Count
    On Error GoTo SafeExit
    
    ' (성능) 표가 없으면: scope 전체에서 Find를 한 번만 수행
    If tableCount = 0 Then
        Dim whole As Range
        Set whole = scopeRange.Duplicate
        Call CollectInSegment(whole, scopeRange, searchText, seen, results)
        Set CollectMatchSpans = results
        Exit Function
    End If
    
    ' 표가 있으면:
    ' 1) 표 밖 구간(테이블 사이 구간)만 큰 덩어리로 Find
    ' 2) 표는 셀 단위로 Find
    Dim curPos As Long
    curPos = scopeS
    
    Dim t As Table
    For Each t In scopeRange.Tables
        Dim ts As Long
        Dim te As Long
        ts = t.Range.Start
        te = t.Range.End
        
        ts = MaxLong(ts, scopeS)
        te = MinLong(te, scopeE)
        
        ' 표 시작 전(표 밖) 구간 수집
        If ts > curPos Then
            Dim outsideSeg As Range
            Set outsideSeg = doc.Range(curPos, ts)
            Call CollectInSegment(outsideSeg, scopeRange, searchText, seen, results)
        End If
        
        ' 표 내부: 셀 단위로 수집
        Dim c As Word.Cell
        For Each c In t.Range.Cells
            ' (중요) 표 셀은 doc.Range(cs, ce)로 다시 만들면 Word가 표 컨텍스트를 잃고
            ' Find가 누락되는 케이스가 있어, Cell.Range.Duplicate를 직접 사용합니다.
            Dim cellR As Range
            Dim cellSeg As Range
            Set cellR = c.Range
            Set cellSeg = cellR.Duplicate
            
            ' 셀 끝 마커 제거:
            ' Cell.Range.Text는 일반적으로 끝에 Chr(13) & Chr(7)가 붙습니다.
            ' Find 안정성을 위해 둘 다 제외(가능한 경우)합니다.
            ' 셀 끝 마커 제거(존재하는 것만 제거)
            ' - 셀 끝에는 항상 Chr(7)이 있으나, 환경/구조에 따라 Chr(13)이 없을 수도 있어
            '   무조건 2글자를 빼면 실제 텍스트가 잘릴 수 있습니다(예: "cbHeigh").
            Do While cellSeg.End > cellSeg.Start
                Dim lastCh As String
                lastCh = cellSeg.Document.Range(cellSeg.End - 1, cellSeg.End).Text
                If lastCh = Chr$(7) Or lastCh = Chr$(13) Then
                    cellSeg.End = cellSeg.End - 1
                Else
                    Exit Do
                End If
            Loop
            
            ' scopeRange로 클램프
            If cellSeg.Start < scopeS Then cellSeg.Start = scopeS
            If cellSeg.End > scopeE Then cellSeg.End = scopeE
            
            If cellSeg.End > cellSeg.Start Then
                ' 표 셀에서는 Find를 쓰지 않고 InStr로 수집(= stuck 근본 제거)
                Call CollectInCellByInStr(cellSeg, scopeRange, searchText, seen, results)
            End If
        Next c
        
        ' 다음 표/표 밖 구간 시작점 갱신
        If te > curPos Then curPos = te
        If curPos >= scopeE Then Exit For
    Next t
    
    ' 마지막 표 뒤(표 밖) 구간 수집
    If curPos < scopeE Then
        Dim tailSeg As Range
        Set tailSeg = doc.Range(curPos, scopeE)
        Call CollectInSegment(tailSeg, scopeRange, searchText, seen, results)
    End If
    
    Set CollectMatchSpans = results
    Exit Function
    
SafeExit:
    ' 실패해도 빈 결과 반환
    If results Is Nothing Then Set results = New Collection
    Set CollectMatchSpans = results
End Function

' 한 구간(seg) 안에서 Find로 매칭(start/end)을 수집합니다.
' - seg는 표 셀 범위 또는 표 밖 문단 범위 등 "작은 단위"여야 합니다.
' - 서식 변경은 절대 하지 않습니다(수집만).
Private Sub CollectInSegment( _
    ByVal seg As Range, _
    ByVal scopeRange As Range, _
    ByVal searchText As String, _
    ByVal seen As Object, _
    ByVal results As Collection _
)
    On Error GoTo SafeExit
    
    If seg Is Nothing Then Exit Sub
    If scopeRange Is Nothing Then Exit Sub
    If searchText = "" Then Exit Sub
    
    Dim doc As Document
    Set doc = seg.Document
    
    Dim endLimit As Long
    endLimit = seg.End
    
    Dim prevS As Long
    Dim prevE As Long
    prevS = -1
    prevE = -1
    
    Dim lastNextPos As Long
    lastNextPos = -1
    
    With seg.Find
        .ClearFormatting
        .Text = searchText
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            Dim ms As Long
            Dim matchEndPos As Long
            ms = seg.Start
            matchEndPos = seg.End
            
            ' 경계 검사(식별자 부분 매치 제거)
            If IsBoundaryMatch(seg, scopeRange) Then
                Dim k As String
                k = CStr(ms) & ":" & CStr(matchEndPos)
                If Not seen.Exists(k) Then
                    seen.Add k, True
                    results.Add Array(ms, matchEndPos)
                End If
            End If
            
            Dim nextPos As Long
            nextPos = matchEndPos
            If nextPos <= ms Then nextPos = ms + 1
            If lastNextPos >= 0 And nextPos <= lastNextPos Then nextPos = lastNextPos + 1

            ' 같은 매치를 반복 반환하는 경우(= stuck)에는 최소 1글자 전진
            If ms = prevS And matchEndPos = prevE Then
                nextPos = ms + 1
                If lastNextPos >= 0 And nextPos <= lastNextPos Then nextPos = lastNextPos + 1
            End If
            
            prevS = ms
            prevE = matchEndPos
            lastNextPos = nextPos
            
            If nextPos >= endLimit Then Exit Do
            seg.SetRange Start:=nextPos, End:=endLimit
        Loop
    End With
    
SafeExit:
End Sub
