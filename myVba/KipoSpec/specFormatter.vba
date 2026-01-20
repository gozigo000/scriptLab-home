Sub ShowNavigationPane()
    ' Navigation pane의 현재 상태를 토글합니다.
    With CommandBars("Navigation")
        .Visible = Not .Visible
    End With
End Sub


Sub ApplyStyleToParagraphs()
    With CommandBars("Navigation")
        .Visible = True
    End With

    Dim doc As Document
    Set doc = ActiveDocument

    With doc.PageSetup
        .TopMargin = CentimetersToPoints(3)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.54)
        .RightMargin = CentimetersToPoints(2.54)
    End With

    With doc.Content.ParagraphFormat
        ' 모든 단락을 양쪽 정렬로 설정
        .Alignment = wdAlignParagraphJustify
        ' 모든 단락의 줄간격을 2줄 간격으로 설정
        .LineSpacingRule = wdLineSpaceDouble
    End With

    Dim para As Paragraph
    ' 제목 1~9 스타일 지정

    Dim targetPara As Paragraph
    Dim wasClaim As Boolean
    Dim claimNumber As Integer

    For Each para In doc.Paragraphs
        If InStr(para.range.Text, "【") > 0 Then
            ' "【"를 포함하는 단락을 찾았을 때 해당 단락의 스타일을 제목 1 스타일로 업데이트
            Set targetPara = para
            para.range.Select
            Exit For
        End If
    Next para

    If Not targetPara Is Nothing Then

        Set selectedRange = Selection.range
        For i = 1 To 9
            doc.Styles("제목 " & i).AutomaticallyUpdate = False
            doc.Styles("제목 " & i).Font = selectedRange.Font
            doc.Styles("제목 " & i).ParagraphFormat = selectedRange.ParagraphFormat
            doc.Styles("제목 " & i).ParagraphFormat.KeepWithNext = True
           
            If i = 1 Then ' "제목 1" 스타일을 가운데 정렬로 설정
                doc.Styles("제목 " & i).ParagraphFormat.Alignment = wdAlignParagraphCenter
            Else
                doc.Styles("제목 " & i).ParagraphFormat.Alignment = wdAlignParagraphJustify
            End If
           
        Next i
    End If


    ' 단락을 찾아서 제목 1~3 스타일로 업데이트

    Dim rng As range
    Set rng = doc.Content

    With rng.Find
        .Text = "【*" ' "【"으로 시작하는 문자열 검색
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True ' 와일드카드 사용 설정
    End With



    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Do While rng.Find.Execute
        ' 찾은 범위에 해당하는 단락에 대한 작업 수행
        'Debug.Print rng.Paragraphs(1).Range.text ' 찾은 단락의 텍스트를 Immediate 창에 출력
        Set para = rng.Paragraphs(1)
       
        Dim paraText As String
        paraText = para.range.Text
       
        If InStr(paraText, "【해결하고자 하는 과제") > 0 Or _
           InStr(paraText, "【과제의 해결 수단") > 0 Or _
           InStr(paraText, "【발명의 효과") > 0 Or _
           InStr(paraText, "【표") > 0 Or _
           InStr(paraText, "【수학식") > 0 Then
            para.style = "제목 3"
        ElseIf InStr(paraText, "【발명의 설명") > 0 Or _
           InStr(paraText, "【청구범위") > 0 Or _
           InStr(paraText, "【요약서") > 0 Or _
           InStr(paraText, "【도면】") > 0 Then
            para.style = "제목 1"
        ElseIf InStr(paraText, "【") > 0 Then
            para.style = "제목 2"
        End If
       
       
        If wasClaim Then
       
            setEnd dict, claimNumber, para
            wasClaim = False
               
        End If
       
       
       
        Dim claimNumMatch As Object
        Set claimNumMatch = CreateObject("VBScript.RegExp")
        claimNumMatch.pattern = "청구항[ ]?([0-9]+)"
        claimNumMatch.Global = False
        If claimNumMatch.TEST(paraText) Then
            Dim matchCollection As Object
            Set matchCollection = claimNumMatch.Execute(paraText)
           
            claimNumber = CInt(matchCollection(0).SubMatches(0))
       
            Dim nextParaStart As Long
            nextParaStart = para.range.End
           
           
            ' 선택된 단락의 인덱스 계산
            Dim paraIndex As Integer
            paraIndex = doc.range(0, para.range.Start).Paragraphs.Count + 1

           
            ' 결과 컬렉션에 추가: 청구항 번호, 다음 단락의 시작 포인트, 숫자 2
            dict.Add claimNumber, Array(2, nextParaStart, -1, paraIndex) ' 청구항 번호 2에 대한 정보
            wasClaim = True
        End If

       
        ' 다음 단락으로 범위 이동
        rng.Collapse Direction:=wdCollapseEnd
        rng.MoveStart wdCharacter, 1
    Loop


    If wasClaim Then
       
            setEnd_Last dict, claimNumber, doc.Paragraphs(doc.Paragraphs.Count)
            wasClaim = False
               
    End If


    Dim key As Variant
    Dim info As Variant
    For Each key In dict.Keys
        info = dict(key)
       ' Debug.Print "청구항 번호: " & key & ", 단락 " & info(3) & ": " & info(1) & "~" & info(2)
       
         Dim paraRange As range
        Set paraRange = doc.range(Start:=info(1), End:=info(2))
       
        Dim upperCLA As Integer
       
        upperCLA = FindSmallestNumberInPattern(paraRange)
        If upperCLA <> 32767 Then
            If dict.Exists(upperCLA) Then
                Dim style As Integer
               
                style = dict(upperCLA)(0) + 1

               
                Dim rngCLA As range
                ' 주어진 종료 포인트에 해당하는 단락을 선택
                Set rngCLA = doc.Paragraphs(info(3)).range
                'rngCLA.Expand Unit:=wdParagraph ' 단락 전체 선택
               
                ' 선택된 단락의 스타일 변경
                rngCLA.style = "제목 " & style
               
                info(0) = style
                dict(key) = info
               
                'rng.Style = newStyle
            End If
        End If
       
        ' 단락을 선택
        'paraRange.Select
       
    Next key


    ChangeTableLineSpacingAndSpacingAfter

End Sub


Sub setEnd(dict As Object, claimNumber As Integer, para As Paragraph)
    If dict.Exists(claimNumber) Then
        Dim findArray As Variant
        findArray = dict(claimNumber)
        ' 정보 처리 : 단락 마지막 포인트 지정
        Dim preParaEnd As Long
        preParaEnd = para.range.Start - 1
                       
        findArray(2) = preParaEnd
        dict(claimNumber) = findArray
    Else

    End If
End Sub

Sub setEnd_Last(dict As Object, claimNumber As Integer, para As Paragraph)
    If dict.Exists(claimNumber) Then
        Dim findArray As Variant
        findArray = dict(claimNumber)
        ' 정보 처리 : 단락 마지막 포인트 지정
        Dim preParaEnd As Long
        preParaEnd = para.range.End
                       
        findArray(2) = preParaEnd
        dict(claimNumber) = findArray
    Else

    End If
End Sub

Function FindSmallestNumberInPattern(rng As range) As Integer
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    Dim matches As Object
    Dim match As Object
    Dim Number As Integer
    Dim smallestNumber As Integer

    ' 정규식 패턴 설정
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "제\s*\d+항|청구항\s*\d+"

    ' 텍스트 범위에서 패턴 찾기
    Set matches = regex.Execute(rng.Text)

    ' 찾은 패턴에서 숫자 추출 및 가장 작은 숫자 찾기
    smallestNumber = 32767 ' 초기값 설정
    For Each match In matches
        Number = Val(Replace(match.Value, "제", "")) ' "제" 제거 후 숫자 추출
        Number = Val(Replace(Number, "청구항", "")) ' "청구항" 제거 후 숫자 추출
        If Number < smallestNumber Then
            smallestNumber = Number
        End If
    Next match

    ' 가장 작은 숫자 반환
    FindSmallestNumberInPattern = smallestNumber
End Function

Sub ChangeTableLineSpacingAndSpacingAfter()
    Dim tbl As Table
    Dim row As row
    Dim cell As cell

    ' 모든 표에 대해 반복
    For Each tbl In ActiveDocument.Tables
        ' 표 안의 모든 셀에 대해 반복
        For Each cell In tbl.range.Cells
                ' 셀 내부의 모든 문단에 대해 반복
                For Each para In cell.range.Paragraphs
                    para.LineSpacingRule = wdLineSpaceSingle ' 줄간격을 1줄(단일)로 설정
                Next para
        Next cell

       
        ' 표 바로 다음 단락에 대해 작업
        With tbl.range.Next(wdParagraph, 1)
            .ParagraphFormat.SpaceBefore = 12 ' 단락 앞 간격을 12pt로 설정
        End With
    Next tbl
    AddSpaceAfterParagraphWithOnlyInlineImage

    'MsgBox "표의 줄간격이 변경되었고, 표 바로 다음 단락에 단락 앞 간격이 추가되었습니다.", vbInformation

End Sub

Sub AddSpaceAfterParagraphWithOnlyInlineImage()
    Dim shape As InlineShape
    Dim para As Paragraph
    Dim rng As range
    Dim paraText As String
    Dim i As Integer

    ' 문서 내의 모든 InlineShape 순회
    For Each shape In ActiveDocument.InlineShapes
        ' InlineShape가 이미지인지 확인
        If shape.Type = wdInlineShapePicture Then
            ' 이미지가 있는 단락 가져오기
            Set para = shape.range.Paragraphs(1)
            ' 단락의 텍스트 가져오기
            paraText = para.range.Text
            ' 텍스트에 공백이 아닌 문자가 있는지 확인
            For i = 1 To Len(paraText)
                If Asc(Mid(paraText, i, 1)) > 32 Then ' ASCII 코드가 32보다 큰 경우 (공백 이후의 문자)
                    ' 단락에 실제 텍스트가 있는 경우
                    Exit For
                End If
            Next i
            ' 텍스트가 없는 경우
            If i > Len(paraText) Then
                shape.range.ParagraphFormat.SpaceAfter = 12
                ' 새로운 Range 객체 생성
                'Set rng = para.Range.Duplicate
                'rng.InsertAfter vbCr & Space(12) ' 12pt 공간 추가
            End If
        End If
    Next shape
End Sub
