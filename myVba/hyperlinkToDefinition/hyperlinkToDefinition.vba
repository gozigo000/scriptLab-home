' (MARK) 선택 영역으로 하이퍼링크 생성 섹션
' ----------------------
' 현재 선택 영역에 있는 단어와 동일한 단어들을 찾아서
' 현재 선택 영역으로 하이퍼링크를 걸어주는 기능
'
' 사용 방법:
' 1. 이 모듈을 Word VBA 프로젝트에 추가합니다.
' 2. 하이퍼링크를 생성하려는 단어를 선택합니다.
' 3. CreateHyperlinksToSelection 매크로를 실행합니다.
'    (예: 매크로 버튼이나 단축키에 할당)

' 선택 영역으로 하이퍼링크 생성
Public Sub CreateHyperlinksToSelection()
    On Error GoTo ErrorHandler
    
    ' 선택 영역이 비어있거나 단어가 아닌 경우 종료
    If Selection.Type = wdSelectionIP Then
        VBA.MsgBox "단어를 선택한 후 실행해주세요.", vbExclamation, "알림"
        Exit Sub
    End If
    
    Dim selectedText As String
    Dim originalRange As Range
    Dim findRange As Range
    Dim bookmarkName As String
    Dim hyperlinkCount As Long
    
    ' 현재 선택 영역 저장
    Set originalRange = Selection.Range.Duplicate
    selectedText = Trim(Selection.Text)
    
    ' 줄바꿈 기호가 있으면 종료 (단일 단어만 지원)
    If InStr(selectedText, vbCrLf) > 0 Or InStr(selectedText, vbLf) > 0 Or InStr(selectedText, vbCr) > 0 Then
        VBA.MsgBox "단일 단어만 선택해주세요.", vbExclamation, "알림"
        Exit Sub
    End If
    
    ' 선택된 텍스트가 비어있으면 종료
    If selectedText = "" Then
        VBA.MsgBox "텍스트를 선택한 후 실행해주세요.", vbExclamation, "알림"
        Exit Sub
    End If
    
    ' 화면 업데이트 일시 중지
    Application.ScreenUpdating = False
    
    ' Undo 스택에 하나의 액션으로 묶기 시작
    Dim undoRecord As UndoRecord
    Set undoRecord = Application.UndoRecord
    undoRecord.StartCustomRecord "하이퍼링크 생성"
    
    ' 현재 선택 영역에 북마크 생성 (하이퍼링크 목적지)
    ' 북마크 이름에 사용할 수 없는 문자 제거 (공백, 특수문자 등)
    Dim safeSelectedText As String
    safeSelectedText = selectedText
    ' 공백을 언더스코어로 변경
    safeSelectedText = Replace(safeSelectedText, " ", "_")
    ' 북마크 이름에 사용할 수 없는 특수문자 제거
    Dim i As Long
    Dim char As String
    Dim validText As String
    validText = ""
    For i = 1 To Len(safeSelectedText)
        char = Mid(safeSelectedText, i, 1)
        ' 영문자, 숫자, 언더스코어만 허용
        If (char >= "A" And char <= "Z") Or (char >= "a" And char <= "z") Or (char >= "0" And char <= "9") Or char = "_" Or (Asc(char) >= 128) Then
            validText = validText & char
        End If
    Next i
    bookmarkName = "HyperlinkTarget_" & validText
    
    ' 기존 북마크가 있으면 삭제
    On Error Resume Next
    ActiveDocument.Bookmarks(bookmarkName).Delete
    On Error GoTo ErrorHandler
    
    ' 현재 선택 영역에 북마크 추가
    originalRange.Bookmarks.Add bookmarkName
    
    ' 원본 선택 영역이 속한 문단의 텍스트 가져오기 (ScreenTip용)
    Dim paragraphText As String
    paragraphText = Trim(originalRange.Paragraphs(1).Range.Text)
    ' 문단 끝의 줄바꿈 문자 제거
    If Right(paragraphText, 1) = vbCr Or Right(paragraphText, 1) = vbLf Or Right(paragraphText, 2) = vbCrLf Then
        paragraphText = Left(paragraphText, Len(paragraphText) - 1)
    End If
    If Right(paragraphText, 1) = vbCr Or Right(paragraphText, 1) = vbLf Then
        paragraphText = Left(paragraphText, Len(paragraphText) - 1)
    End If
    
    ' 선택한 단어를 쌍따옴표로 강조
    If InStr(paragraphText, selectedText) > 0 Then
        paragraphText = Replace(paragraphText, selectedText, """" & selectedText & """")
    End If
    
    ' 문서 전체에서 동일한 텍스트 검색
    Set findRange = ActiveDocument.Content
    hyperlinkCount = 0
    Dim previousStart As Long
    Dim previousEnd As Long
    previousStart = -1
    previousEnd = -1
    
    With findRange.Find
        .ClearFormatting
        .Text = selectedText
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        
        ' 모든 일치 항목에 하이퍼링크 추가
        Do While .Execute
            ' 무한 루프 방지: 같은 위치를 다시 찾으면 종료
            If findRange.Start = previousStart And findRange.End = previousEnd Then
                Exit Do
            End If
            
            ' 현재 위치 저장
            previousStart = findRange.Start
            previousEnd = findRange.End
            
            ' 북마크 영역과 일치하는지 확인
            Dim isBookmarkRange As Boolean
            isBookmarkRange = False
            
            ' 북마크가 존재하는지 확인하고 범위 비교
            On Error Resume Next
            Dim bookmarkRange As Range
            Set bookmarkRange = ActiveDocument.Bookmarks(bookmarkName).Range
            If Err.Number = 0 Then
                ' findRange가 북마크 영역과 겹치는지 확인
                If findRange.Start >= bookmarkRange.Start And findRange.End <= bookmarkRange.End Then
                    isBookmarkRange = True
                End If
            End If
            On Error GoTo ErrorHandler
            
            ' 현재 선택 영역(북마크 위치)은 제외
            If Not isBookmarkRange Then
                ' 하이퍼링크 추가 전에 범위를 복사 (원본 범위가 변경되지 않도록)
                Dim targetRange As Range
                Set targetRange = findRange.Duplicate
                
                ' 하이퍼링크 추가 전에 범위의 끝 위치 저장
                Dim rangeEnd As Long
                rangeEnd = findRange.End
                
                ' 원본 범위의 서식 저장 (전체 서식 참조)
                Dim originalFormatRange As Range
                Set originalFormatRange = findRange.Duplicate
                
                ' 하이퍼링크 추가 전 범위 위치 저장
                Dim targetStart As Long
                Dim targetEnd As Long
                targetStart = targetRange.Start
                targetEnd = targetRange.End
                
                ' 기존 하이퍼링크가 있으면 제거
                If targetRange.Hyperlinks.Count > 0 Then
                    Dim existingHyperlink As Hyperlink
                    For Each existingHyperlink In targetRange.Hyperlinks
                        existingHyperlink.Delete
                    Next existingHyperlink
                End If
                
                ' 하이퍼링크 추가 (같은 문서 내 북마크로 연결)
                Dim newHyperlink As Hyperlink
                Set newHyperlink = ActiveDocument.Hyperlinks.Add( _
                    Anchor:=targetRange, _
                    Address:="", _
                    SubAddress:=bookmarkName, _
                    ScreenTip:=paragraphText, _
                    TextToDisplay:=selectedText)
                
                ' 하이퍼링크 추가 후 하이퍼링크 객체의 Range를 통해 서식 수정
                Dim hyperlinkRange As Range
                Set hyperlinkRange = newHyperlink.Range
                
                ' 원본 범위의 서식을 하이퍼링크 범위에 복원
                ' Font 서식 복원
                hyperlinkRange.Font.Color = originalFormatRange.Font.Color
                hyperlinkRange.Font.Underline = originalFormatRange.Font.Underline
                hyperlinkRange.Font.Bold = originalFormatRange.Font.Bold
                hyperlinkRange.Font.Italic = originalFormatRange.Font.Italic
                hyperlinkRange.Font.Size = originalFormatRange.Font.Size
                hyperlinkRange.Font.Name = originalFormatRange.Font.Name
                
                ' Shading 서식 복원
                hyperlinkRange.Shading.BackgroundPatternColor = originalFormatRange.Shading.BackgroundPatternColor
                hyperlinkRange.Shading.ForegroundPatternColor = originalFormatRange.Shading.ForegroundPatternColor
                
                ' 스타일 복원 (하이퍼링크 스타일이 적용되지 않도록)
                ' On Error Resume Next
                ' hyperlinkRange.Style = originalFormatRange.Style
                ' On Error GoTo ErrorHandler
                
                ' 하이퍼링크 범위 정리
                Set hyperlinkRange = Nothing
                
                ' 원본 범위를 저장한 끝 위치로 설정
                targetRange.SetRange targetStart, targetEnd
                
                hyperlinkCount = hyperlinkCount + 1
                
                ' 정리
                Set targetRange = Nothing
                Set originalFormatRange = Nothing
                
                ' 하이퍼링크 추가 후 원본 범위를 저장한 끝 위치로 설정
                findRange.SetRange rangeEnd, rangeEnd
            End If
            
            ' 다음 검색을 위해 범위를 끝으로 축소
            findRange.Collapse wdCollapseEnd
            
            ' 하이퍼링크 추가 후 위치가 변경되었을 수 있으므로, 
            ' 이전 위치와 비교하여 같은 위치면 한 칸 앞으로 이동
            If findRange.Start = previousEnd Then
                findRange.MoveStart wdCharacter, 1
                findRange.Collapse wdCollapseStart
            End If
            
            ' 문서 끝에 도달했는지 확인
            If findRange.Start >= ActiveDocument.Content.End - 1 Then
                Exit Do
            End If
        Loop
    End With
    
    ' Undo 스택 묶기 종료
    undoRecord.EndCustomRecord
    
    ' 화면 업데이트 재개
    Application.ScreenUpdating = True
    
    ' 원래 선택 영역으로 복원
    originalRange.Select
    
    ' 완료 메시지
    If hyperlinkCount > 0 Then
        VBA.MsgBox selectedText & " 단어에 " & hyperlinkCount & "개의 하이퍼링크를 생성했습니다.", vbInformation, "완료"

        ' 하이퍼링크 스타일 수정 함수 호출 (한 번만 호출되도록 확인)
        On Error Resume Next
        Dim styleModified As Boolean
        styleModified = False
        
        ' 문서 변수 확인: 스타일이 이미 수정되었는지 확인
        Dim varValue As String
        varValue = ActiveDocument.Variables("HyperlinkStyleModified").Value
        If Err.Number = 0 And varValue = "True" Then
            ' 이미 스타일이 수정된 경우
            styleModified = True
        Else
            ' 문서 변수가 없거나 False인 경우, 변수 추가/설정
            Err.Clear
            ActiveDocument.Variables("HyperlinkStyleModified").Value = "True"
            If Err.Number <> 0 Then
                ActiveDocument.Variables.Add Name:="HyperlinkStyleModified", Value:="True"
            End If
        End If
        On Error GoTo ErrorHandler
        
        ' 하이퍼링크 스타일 수정 함수 호출 (아직 수정되지 않은 경우에만)
        If Not styleModified Then
            Call ModifyHyperlinkStylesForLinkingToDefinition
        End If
    Else
        VBA.MsgBox selectedText & " 단어의 다른 일치 항목을 찾을 수 없습니다.", vbInformation, "알림"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Undo 스택 묶기 종료 (오류 발생 시에도)
    On Error Resume Next
    If Not undoRecord Is Nothing Then
        undoRecord.EndCustomRecord
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    VBA.MsgBox "오류가 발생했습니다: " & Err.Description, vbCritical, "오류"
    On Error Resume Next
    If Not originalRange Is Nothing Then
        originalRange.Select
    End If
End Sub

' 기존 하이퍼링크 제거 (선택 영역의 단어와 동일한 단어들의 하이퍼링크 제거)
Public Sub RemoveHyperlinksFromSelection()
    On Error GoTo ErrorHandler
    
    ' 선택 영역이 비어있거나 단어가 아닌 경우 종료
    If Selection.Type = wdSelectionIP Then
        VBA.MsgBox "단어를 선택한 후 실행해주세요.", vbExclamation, "알림"
        Exit Sub
    End If
    
    Dim selectedText As String
    Dim originalRange As Range
    Dim findRange As Range
    Dim hyperlinkCount As Long
    
    ' 현재 선택 영역 저장
    Set originalRange = Selection.Range.Duplicate
    selectedText = Trim(Selection.Text)
    
    ' 줄바꿈 기호가 있으면 종료
    If InStr(selectedText, vbCrLf) > 0 Or InStr(selectedText, vbLf) > 0 Or InStr(selectedText, vbCr) > 0 Then
        VBA.MsgBox "단일 단어만 선택해주세요.", vbExclamation, "알림"
        Exit Sub
    End If
    
    ' 선택된 텍스트가 비어있으면 종료
    If selectedText = "" Then
        VBA.MsgBox "텍스트를 선택한 후 실행해주세요.", vbExclamation, "알림"
        Exit Sub
    End If
    
    ' 화면 업데이트 일시 중지
    Application.ScreenUpdating = False
    
    ' Undo 스택에 하나의 액션으로 묶기 시작
    Dim undoRecord As UndoRecord
    Set undoRecord = Application.UndoRecord
    undoRecord.StartCustomRecord "하이퍼링크 제거"
    
    ' 문서 전체에서 동일한 텍스트 검색
    Set findRange = ActiveDocument.Content
    hyperlinkCount = 0
    
    With findRange.Find
        .ClearFormatting
        .Text = selectedText
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        
        ' 모든 일치 항목의 하이퍼링크 제거
        Do While .Execute
            If findRange.Hyperlinks.Count > 0 Then
                Dim existingHyperlink As Hyperlink
                For Each existingHyperlink In findRange.Hyperlinks
                    existingHyperlink.Delete
                    hyperlinkCount = hyperlinkCount + 1
                Next existingHyperlink
            End If
            
            findRange.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Undo 스택 묶기 종료
    undoRecord.EndCustomRecord
    
    ' 화면 업데이트 재개
    Application.ScreenUpdating = True
    
    ' 원래 선택 영역으로 복원
    originalRange.Select
    
    ' 완료 메시지
    VBA.MsgBox selectedText & " 단어에서 " & hyperlinkCount & "개의 하이퍼링크를 제거했습니다.", vbInformation, "완료"
    
    Exit Sub
    
ErrorHandler:
    ' Undo 스택 묶기 종료 (오류 발생 시에도)
    On Error Resume Next
    If Not undoRecord Is Nothing Then
        undoRecord.EndCustomRecord
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    VBA.MsgBox "오류가 발생했습니다: " & Err.Description, vbCritical, "오류"
    On Error Resume Next
    If Not originalRange Is Nothing Then
        originalRange.Select
    End If
End Sub
