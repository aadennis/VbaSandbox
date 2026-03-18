Option Explicit

Sub ReplaceAnyBulletsWithCheckboxes()

    Dim para As Paragraph
    Dim cc As ContentControl
    Dim rng As Range

    For Each para In ActiveDocument.Paragraphs
        
        If para.Range.ListFormat.ListType <> wdListNoNumbering Then
            
            para.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
            
            'Start-of-paragraph range
            Set rng = para.Range.Duplicate
            rng.Collapse wdCollapseStart
            
            'Insert checkbox
            Set cc = ActiveDocument.ContentControls.Add(wdContentControlCheckBox, rng)
            
            'Move OUTSIDE the content control before inserting spaces
            Set rng = para.Range.Duplicate
            rng.Start = cc.Range.End
            rng.End = rng.Start
            
        End If
    Next para

End Sub

Sub AddSpacesBeforeBulletText()

    Dim para As Paragraph
    Dim rng As Range

    For Each para In ActiveDocument.Paragraphs
        
        'Check if paragraph is part of ANY list (bullets, custom bullets, list styles)
        If para.Range.ListFormat.ListType <> wdListNoNumbering Then
            
            'Create a range at the start of the paragraph text (after the bullet)
            Set rng = para.Range.Duplicate
            rng.Collapse Direction:=wdCollapseStart
            
            'Move to after the bullet/number formatting
            rng.MoveStartWhile Cset:=vbTab & " ", Count:=5
            
            'Insert two spaces before the text
            rng.InsertBefore "  "
        End If
    Next para

End Sub

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
