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
'        Call InitializeSelectionAutoSearch
'        Set myAppEvents = New clsAppEvents
'        Set myAppEvents.appWord = Word.Application
'    End Sub
'
' 5. 기능을 활성화하려면 ToggleSelectionAutoSearch 서브루틴을 호출합니다.
'    예: 매크로 버튼이나 단축키에 ToggleSelectionAutoSearch 할당

' 모듈 레벨 변수 (다른 모듈에서 접근 가능하도록 Public으로 선언)
Public isSelectionAutoSearchEnabled As Boolean
Public previousSelectedText As String
Public isProcessingSelectionChange As Boolean ' 무한루프 방지 플래그

' 초기화 프로시저 (이 모듈이 로드될 때 호출)
Public Sub InitializeSelectionAutoSearch()
    isSelectionAutoSearchEnabled = False
    previousSelectedText = ""
    isProcessingSelectionChange = False
End Sub

' 선택 영역 자동 검색 기능 토글
Public Sub ToggleSelectionAutoSearch()
    isSelectionAutoSearchEnabled = Not isSelectionAutoSearchEnabled
    
    If isSelectionAutoSearchEnabled Then
        ' 기능 활성화
        Call showMsg("선택 영역 자동 검색이 활성화되었습니다.", "알림", vbInformation, 1000)
    Else
        ' 기능 비활성화
        ' 이전 하이라이트 제거
        If previousSelectedText <> "" Then
            Call RemoveHighlight(previousSelectedText)
            previousSelectedText = ""
        End If
        Call showMsg("선택 영역 자동 검색이 비활성화되었습니다.", "알림", vbInformation, 1000)
    End If
End Sub

' 이전 하이라이트 제거 함수
Public Sub RemoveHighlight(searchText As String)
    On Error GoTo ErrorHandler
    
    Dim findRange As Range
    Dim originalRange As Range
    
    ' 현재 커서 위치 저장 (Range 객체로 저장하여 Select 호출 방지)
    Set originalRange = Selection.Range.Duplicate
    
    ' 화면 업데이트 일시 중지 (이벤트 발생 감소)
    Application.ScreenUpdating = False
    
    ' 문서 전체에서 검색
    Set findRange = ActiveDocument.Content
    
    With findRange.Find
        .ClearFormatting
        .Text = searchText
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        
        ' 모든 일치 항목 찾아서 하이라이트 제거
        Do While .Execute
            findRange.Shading.BackgroundPatternColor = wdColorAutomatic ' 자동 색상
            ' findRange.HighlightColorIndex = wdNoHighlight
            findRange.Collapse wdCollapseEnd
        Loop
    End With
    
    ' 화면 업데이트 재개
    Application.ScreenUpdating = True
    
    ' 원래 커서 위치로 복원 (플래그가 설정되어 있으면 Select 호출하지 않음)
    If Not isProcessingSelectionChange Then
        originalRange.Select
    Else
        ' 이벤트 처리 중이면 Range만 이동 (Select 호출 시 무한루프 발생)
        Selection.SetRange originalRange.Start, originalRange.End
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "하이라이트 제거 중 오류: " & Err.Description
    Application.ScreenUpdating = True
    ' 원래 커서 위치로 복원 시도
    On Error Resume Next
    If Not isProcessingSelectionChange Then
        originalRange.Select
    Else
        Selection.SetRange originalRange.Start, originalRange.End
    End If
End Sub

' 선택 영역이 변경될 때 호출되는 함수
' 이 함수는 clsAppEvents 클래스 모듈에서 appWord_WindowSelectionChange 이벤트로 호출됨
Public Sub OnSelectionChanged()
    ' 기능이 비활성화되어 있으면 종료
    If Not isSelectionAutoSearchEnabled Then
        Exit Sub 
    End If
    
    ' 무한루프 방지: 이미 처리 중이면 종료
    If isProcessingSelectionChange Then
        Exit Sub
    End If
    
    ' 처리 중 플래그 설정
    isProcessingSelectionChange = True
    
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    Dim findRange As Range
    Dim originalRange As Range
    
    ' 선택된 텍스트가 없으면 이전 하이라이트 제거하고 종료
    If Selection.Type = wdSelectionIP Then
        If previousSelectedText <> "" Then
            Call RemoveHighlight(previousSelectedText)
            previousSelectedText = ""
        End If
        isProcessingSelectionChange = False
        Exit Sub
    End If

    ' 현재 선택 영역 가져오기
    selectedText = Trim(Selection.Text)
    
    ' 줄바꿈 기호가 있으면 바로 종료
    If InStr(selectedText, vbCrLf) > 0 Or InStr(selectedText, vbLf) > 0 Or InStr(selectedText, vbCr) > 0 Then
        isProcessingSelectionChange = False
        Exit Sub
    End If
    
    ' 이전 텍스트와 동일하면 하이라이트 유지하고 종료
    If selectedText = previousSelectedText Then
        isProcessingSelectionChange = False
        Exit Sub
    End If
    
    ' 현재 커서 위치 저장 (Range 객체로 저장)
    Set originalRange = Selection.Range.Duplicate

    ' 이전에 하이라이트된 텍스트가 있으면 하이라이트 제거
    If previousSelectedText <> "" Then
        Call RemoveHighlight(previousSelectedText)
    End If

    ' 화면 업데이트 일시 중지 (이벤트 발생 감소)
    Application.ScreenUpdating = False

    ' 문서 전체에서 선택된 텍스트와 동일한 텍스트 검색
    Set findRange = ActiveDocument.Content
    
    With findRange.Find
        .ClearFormatting
        .Text = selectedText
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop ' (info) [문서 끝까지 검색 후 동작] wdFindContinue: 처음부터 다시 검색 (무한루프 위험), wdFindAsk: 사용자에게 물어봄, wdFindStop: 검색 중지

        ' 찾은 모든 텍스트에 초록색 하이라이트 적용
        Do While .Execute
            ' Lime 색상 (RGB: 0, 255, 0)을 하이라이트로 적용
            ' Word VBA의 HighlightColorIndex는 제한된 색상만 지원합니다.
            ' Lime에 가장 가까운 색상은 wdTurquoise입니다.
            ' 정확한 Lime 색상을 원하면 아래 주석을 해제하고 위 줄을 주석 처리하세요:
            findRange.Shading.BackgroundPatternColor = RGB(0, 255, 0) ' Lime 색상
            ' findRange.HighlightColorIndex = wdNoHighlight ' 하이라이트 제거
            ' findRange.HighlightColorIndex = wdTurquoise ' Lime에 가까운 색상
            findRange.Collapse wdCollapseEnd ' 검색범위의 시작을 검색된 텍스트의 끝으로 이동
        Loop
    End With
    
    ' 화면 업데이트 재개
    Application.ScreenUpdating = True
    
    ' 원래 커서 위치로 복원 (SetRange 사용하여 Select 호출 방지)
    Selection.SetRange originalRange.Start, originalRange.End
    
    ' 현재 선택된 텍스트를 이전 텍스트로 저장
    previousSelectedText = selectedText
    
    isProcessingSelectionChange = False
    Exit Sub
    
ErrorHandler:
    Debug.Print "선택 영역 검색 및 하이라이트 적용 중 오류: " & Err.Description
    Application.ScreenUpdating = True
    ' 원래 커서 위치로 복원 시도
    On Error Resume Next
    If Not originalRange Is Nothing Then
        Selection.SetRange originalRange.Start, originalRange.End
    End If
End Sub

