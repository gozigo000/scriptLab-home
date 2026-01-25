Option Explicit

' =============================================================================
' Module: specFormatter
' Purpose:
'   - KIPO spec 문서에서 "【...】" 단락 규칙과 청구항 참조 관계를 이용해
'     Word 스타일(제목 1~9) 및 표/이미지 주변 문단 간격을 자동 정리합니다.
'
' Public API:
'   - formatKipoSpec()
' =============================================================================
' TODO:
' 
' [KIPO포매터]
' 저장할 때 자동으로 필드 업데이트하기
' 독립항들은 배경색 칠하기
' 청구범위 청구항 하이퍼링크
' 도 #a 하이퍼링크 (도 1, 도 11 구분 주의)
' 도면들 들여쓰기 제거
' 볼드체 문단 -> 제목 2 & 들여쓰기 유지 & 위에 한 줄 추가
' 상용구 배경 어둡게 색칠: 완전 일치 => 어둡게  /  부분 일치 => 덜 어둡게
' 
' CreateHyperlinksToSelection 함수가 Range 인수를 받도록 수정
' clsAppEvents 객체 유지하기
' 문서 열었을 때 개요 접기 상태 세팅
' 
' [하이라이트 고려대상]
' C, TL, TR, BL, BR, L', T', BL', TR', TL' , L0, L1, L2
' non-CCP, SAD, IBC, 
' N
' rec'L
' HEVC, VVC, ECM, ECM-11.0
' 아래첨자: 빨간색 칠하기
' 
' [US OA 포매터]
' 날짜 색칠: March 16, 2013  /  08/20/2025
' 법조문 색칠: 35 U.S.C. 102  /  35 U.S.C. 102 and 103  /  35 U.S.C. 102(a)(1)
' 인용문헌 색칠: D1, D2, ... /  Zhang et al.  /  
' 심사관 색칠: The Examiner / the examiner 
' 출원인 색칠: The applicant
' 청구항 색칠: claims #-##  /  claim #  /  Claims 1, 8, 10, 14, 16  /  
' 도면 색칠: Figure 6  /  Fig. 11  /  
' 문단 색칠: [0032]
' 이탤릭체(인용부분) 색칠: 
' have become moot (거절이유 치유) 색칠
' =============================================================================

' 실행 매크로.
' - "【...】" 단락에 제목 스타일(1~3)을 적용합니다.
' - 청구항 참조(제 n항/청구항 n)를 기반으로 제목 레벨을 추정해 조정합니다.
Public Sub formatKipoSpec()
    On Error GoTo ErrorHandler
    Call BeginCustomUndoRecord("formatKipoSpec")
    
    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim doc As Document
    Set doc = ActiveDocument

    Call SetPageFormat(doc)
    Call SetNormalStyleDefaults(doc)
    Call SetDocumentParagraphDefaults(doc)
    Call SetHeadingStyles(doc)
    Call EraseDirectFormatting(doc)

    Call InsertBlankLinesForKipoSections(doc)
    Call InsertBlankLinesBetweenClaimHeadings(doc)
    Call InsertBlankLinesBeforeFigureHeadings(doc)
    
    Call ApplyHeadingStyles(doc)
    Call AdjustClaimHeadingLevels(doc)

    Call ColorizeText(doc)
    Call ShadeText(doc)

    Call ShowNavigationPane()
    GoTo SafeExit

ErrorHandler:
    VBA.MsgBox "오류: " & Err.Description, vbCritical, "specFormatter"
    GoTo SafeExit

SafeExit:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    Call EndCustomUndoRecord()
End Sub

' 영문 대문자/소문자/숫자를 각각 다른 색으로 표시합니다.
' - 대문자: 진한 파랑
' - 소문자: 진한 초록
' - 숫자: 진한 빨강
Private Sub ColorizeText(ByVal doc As Document)
    On Error GoTo SafeExit
    ' Call ApplyColorByWildcard(doc, "[a-zA-Z_]@", RGB(237, 125, 49), True)
    ' Call ApplyColorByWildcard(doc, "[a-z]@", RGB(0, 153, 74), True)
    ' Call ApplyColorByWildcard(doc, "[0-9]@", RGB(204, 0, 0), False)
SafeExit:
End Sub

' 문서 전체에서 "도 1", "도2" 같은 도면 참조 토큰의 배경색을 지정합니다.
' - 요청 배경색: RGB(255,242,204)
Private Sub ShadeText(ByVal doc As Document)
    On Error GoTo SafeExit

    Dim figureBgColor As Long
    figureBgColor = RGB(255, 242, 204)

    Dim contentRng As Range
    Set contentRng = doc.Content

    Dim s As String
    s = contentRng.Text

    Dim matches As Object
    Set matches = GetRegexMatchesAll(s, "(도|표|수학식)\s*\d+[a-zA-Z]?")
    If matches Is Nothing Then GoTo SafeExit

    Dim m As Object
    For Each m In matches
        Dim startIdx As Long
        Dim tokenLen As Long
        startIdx = CLng(m.FirstIndex) ' 0-based
        tokenLen = Len(CStr(m.Value))
        If (startIdx >= 1 And _
            IsDelimiterChar(Mid$(s, startIdx, 1)) _
        ) Then
            Dim hit As Range
            Set hit = doc.Range( _
                contentRng.Start + startIdx, _
                contentRng.Start + startIdx + tokenLen _
            )

            hit.Shading.Texture = wdTextureNone
            hit.Shading.BackgroundPatternColor = figureBgColor
        End If
    Next m

SafeExit:
End Sub

Private Sub ApplyColorByWildcard( _
    ByVal doc As Document, _
    ByVal pattern As String, _
    ByVal rgbColor As Long, _
    ByVal matchCase As Boolean _
)
    On Error GoTo SafeExit

    Dim rng As Range
    Set rng = doc.Content

    With rng.Find
        .ClearFormatting
        .Text = pattern
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        .MatchCase = matchCase
    End With

    Do While rng.Find.Execute
        rng.Font.Color = rgbColor

        rng.Collapse Direction:=wdCollapseEnd
    Loop

SafeExit:
End Sub

' 페이지 포맷 설정
Private Sub SetPageFormat(ByVal doc As Document)
    ' 페이지 여백을 설정합니다.
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(3)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.54)
        .RightMargin = CentimetersToPoints(2.54)
    End With
End Sub

' =============================================================================
' Normal(기본) 스타일 기본값
' =============================================================================

' Normal(기본) 스타일의 폰트/단락 서식을 한 곳에서 설정합니다.
Private Sub SetNormalStyleDefaults(ByVal doc As Document)
    On Error GoTo SafeExit

    Const DEFAULT_FONT_SIZE As Single = 12
    Const DEFAULT_FIRST_LINE_INDENT_CM As Single = 1.41
    Const DEFAULT_LINE_SPACING_MULTIPLE As Single = 1.6

    ' "기본 단락 스타일"은 Normal(기본) 스타일을 의미합니다.
    ' doc.Content.ParagraphFormat에 들여쓰기를 주면 제목 스타일에도 직접 서식으로
    ' 적용될 수 있어, Normal 스타일에만 들여쓰기/폰트 크기를 적용합니다.
    With doc.Styles(wdStyleNormal).Font
        .NameFarEast = "맑은 고딕"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = DEFAULT_FONT_SIZE
    End With

    With doc.Styles(wdStyleNormal).ParagraphFormat
        .Alignment = wdAlignParagraphJustify
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(DEFAULT_LINE_SPACING_MULTIPLE)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = CentimetersToPoints(DEFAULT_FIRST_LINE_INDENT_CM)
        .LeftIndent = 0
        .RightIndent = 0
    End With

SafeExit:
End Sub

' 문서 전체 기본 단락 서식(직접 서식)을 설정합니다.
' - 제목/기타 스타일까지 영향이 갈 수 있으므로, 들여쓰기는 여기서 건드리지 않습니다.
Private Sub SetDocumentParagraphDefaults(ByVal doc As Document)
    On Error GoTo SafeExit

    Dim baseParaFmt As ParagraphFormat
    Set baseParaFmt = doc.Styles(wdStyleNormal).ParagraphFormat

    With doc.Content.ParagraphFormat
        .Alignment = baseParaFmt.Alignment
        .LineSpacingRule = baseParaFmt.LineSpacingRule
        .LineSpacing = baseParaFmt.LineSpacing
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With

SafeExit:
End Sub

' 제목(1-9) 스타일 설정
' - Normal(기본) 스타일과 동일하게 맞추되, 들여쓰기만 제거합니다.
Private Sub SetHeadingStyles(ByVal doc As Document)
    On Error GoTo SafeExit

    Dim baseFont As Font
    Dim baseParaFmt As ParagraphFormat

    Set baseFont = doc.Styles(wdStyleNormal).Font
    Set baseParaFmt = doc.Styles(wdStyleNormal).ParagraphFormat

    Dim i As Long
    For i = 1 To 9
        With doc.Styles("제목 " & CStr(i))
            .AutomaticallyUpdate = False
            .Font = baseFont

            ' Heading 의미(OutlineLevel 등)는 유지하고, "보이는 서식"만 Normal과 동일하게
            ' 맞춥니다. (줄간격/문단 전후 간격/정렬 등)
            With .ParagraphFormat
                .LineSpacingRule = baseParaFmt.LineSpacingRule
                .LineSpacing = baseParaFmt.LineSpacing
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 0
                .SpaceAfter = 0

                If i = 1 Then
                    .Alignment = wdAlignParagraphCenter
                Else
                    .Alignment = wdAlignParagraphJustify
                End If

                ' 들여쓰기만 제거(첫 줄/좌/우)
                .FirstLineIndent = 0
                .LeftIndent = 0
                .RightIndent = 0

                .KeepWithNext = True
            End With
        End With
    Next i

SafeExit:
End Sub

' 문서 전체에서 "직접 서식"을 제거하고, 각 문단이 가진 스타일 서식을 다시 적용합니다.
' - 본문 폰트가 Normal 변경을 따라오지 않는 경우(직접 서식 잔존) 해결용
Private Sub EraseDirectFormatting(ByVal doc As Document)
    On Error GoTo SafeExit

    Dim rng As Range
    Set rng = doc.Content

    rng.Font.Reset
    rng.ParagraphFormat.Reset

SafeExit:
End Sub

Private Sub ApplyHeadingStyles(ByVal doc As Document)
    Dim rng As Range
    Set rng = doc.Content

    With rng.Find
        .Text = "【*"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
    End With

    Do While rng.Find.Execute
        Dim para As Paragraph
        Set para = rng.Paragraphs(1)
        Call ApplyHeadingStyleToBracketParagraph(para)

        rng.Collapse Direction:=wdCollapseEnd
        rng.MoveStart wdCharacter, 1
    Loop
End Sub

Private Sub ApplyHeadingStyleToBracketParagraph(ByVal para As Paragraph)
    Dim paraText As String
    paraText = para.Range.Text

    ' 제목 3
    If InStr(paraText, "【해결하고자 하는 과제") > 0 Or _
       InStr(paraText, "【기술적 과제") > 0 Or _
       InStr(paraText, "【과제의 해결 수단") > 0 Or _
       InStr(paraText, "【기술적 해결방법") > 0 Or _
       InStr(paraText, "【발명의 효과") > 0 Or _
       InStr(paraText, "【표") > 0 Or _
       InStr(paraText, "【수학식") > 0 Then
        para.Style = "제목 3"
        Exit Sub
    End If

    ' 제목 1
    If InStr(paraText, "【발명의 설명") > 0 Or _
       InStr(paraText, "【명세서") > 0 Or _
       InStr(paraText, "【청구범위") > 0 Or _
       InStr(paraText, "【청구의 범위") > 0 Or _
       InStr(paraText, "【요약서") > 0 Or _
       InStr(paraText, "【도면】") > 0 Then
        para.Style = "제목 1"
        Exit Sub
    End If

    ' 제목 2
    If InStr(paraText, "【") > 0 Then
        para.Style = "제목 2"
    End If
End Sub

Private Sub InsertBlankLinesForKipoSections(ByVal doc As Document)
    On Error GoTo SafeExit

    Dim i As Long
    For i = doc.Paragraphs.Count To 1 Step -1
        Dim para As Paragraph
        Set para = doc.Paragraphs(i)

        If ShouldInsertBlankLineAbove(para.Range.Text) Then
            Call EnsureOneBlankParagraphBefore(doc, para)
        End If
    Next i

SafeExit:
End Sub

Private Sub InsertBlankLinesBetweenClaimHeadings(ByVal doc As Document)
    On Error GoTo SafeExit

    Dim i As Long
    For i = doc.Paragraphs.Count To 2 Step -1
        Dim para As Paragraph
        Dim prevPara As Paragraph

        Set para = doc.Paragraphs(i)
        If Not IsClaimHeadingParagraphText(para.Range.Text) Then GoTo ContinueLoop

        If Not HasPreviousClaimHeading(doc, i - 1) Then GoTo ContinueLoop

        Set prevPara = doc.Paragraphs(i - 1)
        If TextHasOnlyWhitespace(prevPara.Range.Text) Then GoTo ContinueLoop

        Call EnsureOneBlankParagraphBefore(doc, para)

ContinueLoop:
    Next i

SafeExit:
End Sub

Private Sub InsertBlankLinesBeforeFigureHeadings(ByVal doc As Document)
    On Error GoTo SafeExit

    Dim i As Long
    For i = doc.Paragraphs.Count To 2 Step -1
        Dim para As Paragraph
        Dim prevPara As Paragraph

        Set para = doc.Paragraphs(i)
        If Not IsFigureHeadingParagraphText(para.Range.Text) Then GoTo ContinueLoop

        Set prevPara = doc.Paragraphs(i - 1)
        If TextHasOnlyWhitespace(prevPara.Range.Text) Then GoTo ContinueLoop

        Call EnsureOneBlankParagraphBefore(doc, para)

ContinueLoop:
    Next i

SafeExit:
End Sub

Private Function HasPreviousClaimHeading(ByVal doc As Document, ByVal fromIndex As Long) As Boolean
    Dim i As Long
    For i = fromIndex To 1 Step -1
        If IsClaimHeadingParagraphText(doc.Paragraphs(i).Range.Text) Then
            HasPreviousClaimHeading = True
            Exit Function
        End If
    Next i
    HasPreviousClaimHeading = False
End Function

Private Function IsClaimHeadingParagraphText(ByVal paraText As String) As Boolean
    ' ex) 【청구항 1】, 【청구항1】, 【청구항   12】 ...
    IsClaimHeadingParagraphText = TestRegex(paraText, "^\s*【청구항\s*\d+】")
End Function

Private Function IsFigureHeadingParagraphText(ByVal paraText As String) As Boolean
    ' ex) 【도 1】, 【도1】, 【도   12】 ...
    IsFigureHeadingParagraphText = TestRegex(paraText, "^\s*【도\s*\d+】")
End Function

Private Function ShouldInsertBlankLineAbove(ByVal paraText As String) As Boolean
    ShouldInsertBlankLineAbove = _
        InStr(paraText, "【발명의 상세한 설명") > 0 Or _
        InStr(paraText, "【도면의 간단한 설명") > 0 Or _
        InStr(paraText, "【발명의 실시를 위한 형태") > 0
End Function

Private Sub EnsureOneBlankParagraphBefore(ByVal doc As Document, ByVal para As Paragraph)
    On Error GoTo SafeExit

    If IsFirstParagraph(doc, para) Then
        Dim startIp As Range
        Set startIp = doc.Range(0, 0)
        startIp.InsertBefore vbCr
        Exit Sub
    End If

    Dim prevPara As Paragraph
    Set prevPara = GetPreviousParagraph(doc, para)
    If prevPara Is Nothing Then GoTo SafeExit

    If TextHasOnlyWhitespace(prevPara.Range.Text) Then Exit Sub

    Dim ip As Range
    Set ip = doc.Range(para.Range.Start, para.Range.Start)
    ip.InsertBefore vbCr

SafeExit:
End Sub

Private Function IsFirstParagraph(ByVal doc As Document, ByVal para As Paragraph) As Boolean
    On Error GoTo SafeExit
    IsFirstParagraph = (doc.Range(0, para.Range.Start).Paragraphs.Count <= 1)
    Exit Function
SafeExit:
    IsFirstParagraph = False
End Function

Private Function GetPreviousParagraph(ByVal doc As Document, ByVal para As Paragraph) As Paragraph
    On Error GoTo SafeExit
    If para.Range.Start <= 0 Then GoTo SafeExit

    Dim r As Range
    Set r = doc.Range(para.Range.Start - 1, para.Range.Start - 1)
    Set GetPreviousParagraph = r.Paragraphs(1)
    Exit Function

SafeExit:
    Set GetPreviousParagraph = Nothing
End Function

Private Function TextHasOnlyWhitespace(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Asc(Mid$(s, i, 1)) > 32 Then
            TextHasOnlyWhitespace = False
            Exit Function
        End If
    Next i
    TextHasOnlyWhitespace = True
End Function

Private Sub AdjustClaimHeadingLevels(ByVal doc As Document)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Call BuildClaimRangeIndex(doc, dict)
    If dict.Count = 0 Then Exit Sub

    Dim key As Variant
    For Each key In dict.Keys
        Call UpdateClaimHeadingLevel(doc, dict, CLng(key))
    Next key
End Sub

Private Sub BuildClaimRangeIndex(ByVal doc As Document, ByVal dict As Object)
    Dim rng As Range
    Set rng = doc.Content

    With rng.Find
        .Text = "【*"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
    End With

    Dim wasClaim As Boolean
    wasClaim = False

    Dim lastClaimNo As Long
    lastClaimNo = 0

    Do While rng.Find.Execute
        Dim para As Paragraph
        Set para = rng.Paragraphs(1)

        Dim paraText As String
        paraText = para.Range.Text

        If wasClaim Then
            Call SetClaimEnd(dict, lastClaimNo, para.Range.Start - 1)
            wasClaim = False
        End If

        Dim claimNo As Long
        If TryExtractClaimNumber(paraText, claimNo) Then
            Dim nextParaStart As Long
            nextParaStart = para.Range.End

            Dim paraIndex As Long
            paraIndex = doc.Range(0, para.Range.Start).Paragraphs.Count + 1

            dict.Add claimNo, Array(2, nextParaStart, -1, paraIndex)
            wasClaim = True
            lastClaimNo = claimNo
        End If

        rng.Collapse Direction:=wdCollapseEnd
        rng.MoveStart wdCharacter, 1
    Loop

    If wasClaim Then
        Call SetClaimEnd( _
            dict, _
            lastClaimNo, _
            doc.Paragraphs(doc.Paragraphs.Count).Range.End _
        )
    End If
End Sub

Private Sub SetClaimEnd(ByVal dict As Object, ByVal claimNo As Long, ByVal endPos As Long)
    If Not dict.Exists(claimNo) Then Exit Sub

    Dim info As Variant
    info = dict(claimNo)
    info(2) = endPos
    dict(claimNo) = info
End Sub

Private Function TryExtractClaimNumber(ByVal text As String, ByRef outNo As Long) As Boolean
    On Error GoTo SafeExit

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "청구항[ ]?([0-9]+)"
    re.Global = False
    re.IgnoreCase = False

    If Not re.Test(text) Then GoTo SafeExit

    Dim matches As Object
    Set matches = re.Execute(text)

    outNo = CLng(matches(0).SubMatches(0))
    TryExtractClaimNumber = True
    Exit Function

SafeExit:
    outNo = 0
    TryExtractClaimNumber = False
End Function

Private Sub UpdateClaimHeadingLevel( _
    ByVal doc As Document, _
    ByVal dict As Object, _
    ByVal claimNo As Long _
)
    Dim info As Variant
    info = dict(claimNo)

    Dim startPos As Long
    startPos = CLng(info(1))

    Dim endPos As Long
    endPos = CLng(info(2))
    If endPos < startPos Then Exit Sub

    Dim paraRng As Range
    Set paraRng = doc.Range(startPos, endPos)

    Dim upperClaimNo As Long
    upperClaimNo = FindSmallestClaimReferenceNo(paraRng)

    If upperClaimNo <= 0 Then Exit Sub
    If Not dict.Exists(upperClaimNo) Then Exit Sub

    Dim newLevel As Long
    newLevel = CLng(dict(upperClaimNo)(0)) + 1

    Dim claimParaIdx As Long
    claimParaIdx = CLng(info(3))

    Dim claimRng As Range
    Set claimRng = doc.Paragraphs(claimParaIdx).Range
    claimRng.Style = "제목 " & CStr(newLevel)

    info(0) = newLevel
    dict(claimNo) = info
End Sub

Private Function FindSmallestClaimReferenceNo(ByVal rng As Range) As Long
    On Error GoTo SafeExit

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "제\s*\d+항|청구항\s*\d+"

    Dim matches As Object
    Set matches = re.Execute(rng.Text)

    Dim smallest As Long
    smallest = 0

    Dim m As Object
    For Each m In matches
        Dim n As Long
        n = ExtractFirstLong(CStr(m.Value))
        If n <= 0 Then GoTo ContinueLoop

        If smallest = 0 Or n < smallest Then
            smallest = n
        End If

ContinueLoop:
    Next m

    FindSmallestClaimReferenceNo = smallest
    Exit Function

SafeExit:
    FindSmallestClaimReferenceNo = 0
End Function

Private Function ExtractFirstLong(ByVal s As String) As Long
    On Error GoTo SafeExit

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(\d+)"

    If Not re.Test(s) Then GoTo SafeExit

    Dim matches As Object
    Set matches = re.Execute(s)
    ExtractFirstLong = CLng(matches(0).SubMatches(0))
    Exit Function

SafeExit:
    ExtractFirstLong = 0
End Function


Private Sub ShowNavigationPane()
    On Error Resume Next
    With CommandBars("Navigation")
        .Visible = True
    End With
    On Error GoTo 0
End Sub
