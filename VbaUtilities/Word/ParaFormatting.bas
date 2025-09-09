Sub FormatCondolenceName()
    ' C/S C
    With Selection.Paragraphs(1).Range
        ' Font settings
        .Font.Name = "Raleway Medium"
        .Font.Size = 26

        ' Paragraph alignment
        .ParagraphFormat.Alignment = wdAlignParagraphRight
    End With
End Sub
Sub InsertPageBreakAfterBorderedParagraph()
    ' C/S I
    Dim para As Paragraph
    Set para = Selection.Paragraphs(1)

    ' Move cursor to end of the paragraph
    para.Range.Collapse Direction:=wdCollapseEnd

    ' Insert page break
    para.Range.InsertBreak Type:=wdPageBreak
End Sub
