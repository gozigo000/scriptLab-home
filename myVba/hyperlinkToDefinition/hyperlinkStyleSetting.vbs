' 하이퍼링크 스타일 수정: 클릭해도 색상이 변경되지 않도록 설정
Public Sub ModifyHyperlinkStylesForLinkingToDefinition()
    On Error Resume Next
    
    ' Hyperlink(하이퍼링크) 스타일 가져오기
    Dim hyperlinkStyle As Style
    Set hyperlinkStyle = ActiveDocument.Styles("Hyperlink")
    If hyperlinkStyle Is Nothing Then
        Set hyperlinkStyle = ActiveDocument.Styles("하이퍼링크")
    End If
    
    ' FollowedHyperlink(열어본 하이퍼링크) 스타일 가져오기 
    Dim followedHyperlinkStyle As Style
    Set followedHyperlinkStyle = ActiveDocument.Styles("FollowedHyperlink")
    If followedHyperlinkStyle Is Nothing Then
        Set followedHyperlinkStyle = ActiveDocument.Styles("열어본 하이퍼링크")
    End If
    
    ' Hyperlink 스타일 수정
    If Not hyperlinkStyle Is Nothing Then
        ' 하이퍼링크 스타일의 밑줄 제거
        hyperlinkStyle.Font.Underline = wdUnderlineNone
        ' 하이퍼링크 스타일의 색상을 자동(검은색)으로 설정
        hyperlinkStyle.Font.Color = wdColorAutomatic
    End If
    
    ' FollowedHyperlink 스타일 수정
    If Not followedHyperlinkStyle Is Nothing Then
        ' 방문한 하이퍼링크의 색상을 Hyperlink와 동일하게 설정
        ' 이렇게 하면 클릭해도 색상이 변경되지 않음
        If Not hyperlinkStyle Is Nothing Then
            followedHyperlinkStyle.Font.Color = hyperlinkStyle.Font.Color
            followedHyperlinkStyle.Font.Underline = hyperlinkStyle.Font.Underline
        Else
            ' Hyperlink 스타일이 없는 경우 자동 색상으로 설정
            followedHyperlinkStyle.Font.Color = wdColorAutomatic
            followedHyperlinkStyle.Font.Underline = wdUnderlineNone
        End If
    End If
    
    On Error GoTo 0
End Sub
