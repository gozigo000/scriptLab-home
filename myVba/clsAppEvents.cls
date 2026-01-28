' 클래스 모듈: clsAppEvents
' ----------------------
' Application 이벤트를 처리하기 위한 클래스 모듈입니다.
' 
' 이 파일을 VBA 프로젝트에 추가하려면:
' 1. VBA 편집기에서 "삽입" > "클래스 모듈" 선택
' 2. 속성 창에서 Name을 "clsAppEvents"로 변경
' 3. 이 파일의 내용을 복사하여 붙여넣기

Public WithEvents appWord As Word.Application

' (MARK) WindowSelectionChange 폭주 방지(중요)
' Word VBA 이벤트는 단일 스레드(UI 스레드)에서 동작합니다.
' 따라서 아래 3개 호출은 "병렬"이 아니라 "순차"로 실행되며,
' 처리 중에는 UI가 막혀 Windows가 '응답 없음'으로 보일 수 있습니다.
'
' 해결: (1) 재진입 가드, (2) 동일 위치 중복 호출 스킵, (3) 짧은 디바운스(스로틀)
Private Const SELECTION_CHANGE_DEBOUNCE_SEC As Double = 0.08 ' 80ms

' (MARK) SelectionChange 성능 프로파일링(디버깅용)
' - True면, WindowSelectionChange에서 호출별 소요 시간을 Immediate Window에 출력합니다.
' - "새 섹션/표"에서 랙이 걸릴 때 어떤 함수가 병목인지 쉽게 확인하려는 용도입니다.
Private Const PROFILE_SELECTION_CHANGE As Boolean = True
Private Const PROFILE_THRESHOLD_MS As Double = 40 ' 이 시간(ms) 넘으면 출력

' 선택 영역이 변경될 때 발생하는 이벤트
Private Sub appWord_WindowSelectionChange(ByVal Sel As Selection)

    On Error GoTo SafeExit

    Static sIsHandling As Boolean
    Static sLastDocFullName As String
    Static sLastPos As Long
    Static sLastTick As Double

    If sIsHandling Then Exit Sub

    ' 텍스트 드래그 선택은 heavy 기능들에 불리하니 스킵(각 모듈에서도 체크하지만 여기서 1차 차단)
    If Sel.Range.Start <> Sel.Range.End Then Exit Sub

    ' 동일 위치 중복 이벤트 스킵
    Dim docName As String
    On Error Resume Next
    docName = Sel.Range.Document.FullName
    If docName = "" Then docName = Sel.Range.Document.Name
    On Error GoTo SafeExit

    If docName = sLastDocFullName And Sel.Range.Start = sLastPos Then Exit Sub

    ' 디바운스/스로틀: 너무 잦은 호출은 합쳐서 처리
    Dim nowTick As Double
    nowTick = Timer
    If nowTick < sLastTick Then sLastTick = 0 ' 자정 롤오버 보정
    If (nowTick - sLastTick) < SELECTION_CHANGE_DEBOUNCE_SEC Then Exit Sub

    sIsHandling = True
    sLastTick = nowTick
    sLastDocFullName = docName
    sLastPos = Sel.Range.Start

    Dim t0 As Double
    Dim dt As Double

    ' bracketMatcher 모듈의 함수 호출 (괄호 매칭 및 하이라이트)
    t0 = Timer
    Call HighlightBracket(Sel.Range)
    dt = (Timer - t0) * 1000#
    If PROFILE_SELECTION_CHANGE And dt >= PROFILE_THRESHOLD_MS Then
        Debug.Print "WindowSelectionChange: HighlightBracket = " & Format$(dt, "0.0") & "ms"
    End If

    ' currWordHighlighter 모듈의 함수 호출 (현재 단어 하이라이트)
    t0 = Timer
    Call HighlightCurrWord(Sel.Range)
    dt = (Timer - t0) * 1000#
    If PROFILE_SELECTION_CHANGE And dt >= PROFILE_THRESHOLD_MS Then
        Debug.Print "WindowSelectionChange: HighlightCurrWord = " & Format$(dt, "0.0") & "ms"
    End If
    
    ' cursorTracker 모듈의 함수 호출 (문서별 커서 이동 히스토리 저장)
    t0 = Timer
    Call TrackCursorMove(Sel.Range)
    dt = (Timer - t0) * 1000#
    If PROFILE_SELECTION_CHANGE And dt >= PROFILE_THRESHOLD_MS Then
        Debug.Print "WindowSelectionChange: TrackCursorMove = " & Format$(dt, "0.0") & "ms"
    End If

    ' (MARK) 자동 저장(디바운스)
    Call MaybeScheduleAutoSave(Sel.Range.Document)

SafeExit:
    sIsHandling = False
End Sub

Private Sub appWord_DocumentBeforeClose(ByVal Doc As Document, ByRef Cancel As Boolean)
    On Error GoTo SafeExit

    Call CancelAutoSave(Doc)

    ' 하이퍼링크 이동(Alt+→)에서 남긴 북마크 정리
    Call ClearNavBackStack(Doc)
    Call ClearNavForwardStack(Doc)

SafeExit:
End Sub

Private Sub appWord_DocumentBeforeSave( _
    ByVal Doc As Document, _
    ByRef SaveAsUI As Boolean, _
    ByRef Cancel As Boolean _
)
    On Error GoTo SafeExit

    Static sIsHandling As Boolean
    If sIsHandling Then Exit Sub

    sIsHandling = True

    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo SafeExit

    ' 모든 필드 업데이트
    Call UpdateAllFieldsInDocument(Doc)

SafeExit:
    Application.ScreenUpdating = prevScreenUpdating
    sIsHandling = False
End Sub

Private Sub UpdateAllFieldsInDocument(ByVal doc As Word.Document)
    On Error GoTo SafeExit

    Dim story As Word.Range
    For Each story In doc.StoryRanges
        Call UpdateFieldsInStory(story)
    Next story

SafeExit:
End Sub

Private Sub UpdateFieldsInStory(ByVal rng As Word.Range)
    On Error GoTo SafeExit

    Dim cur As Word.Range
    Set cur = rng

    Do While Not cur Is Nothing
        On Error Resume Next
        cur.Fields.Update
        On Error GoTo SafeExit
        Set cur = cur.NextStoryRange
    Loop

SafeExit:
End Sub

