    Sub ResizeAndCenterImages()
    
    Const MAX_WIDTH_CM As Single = 16
    Const MAX_HEIGHT_CM As Single = 20
    For Each pic In ActiveDocument.InlineShapes
        With pic
            If .Type = wdInlineShapePicture Then
                .LockAspectRatio = msoTrue
                If .Width > CentimetersToPoints(MAX_WIDTH_CM) Then
                    .Width = CentimetersToPoints(MAX_WIDTH_CM)
                End If
                If .Height > CentimetersToPoints(MAX_HEIGHT_CM) Then
                    .Height = CentimetersToPoints(MAX_HEIGHT_CM)
                End If
                .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            End If
        End With
    Next pic