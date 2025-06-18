Sub DeleteAllHorizontalLines()
' Delete all horizontal lines in the active document.
' This is to target migrations from Google Docs to Word,
' where horizontal lines are often inserted as inline shapes, and look
' weird in Word.
    Dim para As Paragraph

    For Each para In ActiveDocument.Paragraphs
        ' Look for horizontal line elements (field codes or special styles)
        If para.Range.InlineShapes.Count > 0 Then
            Dim ils As InlineShape
            For Each ils In para.Range.InlineShapes
                If ils.Type = wdInlineShapeHorizontalLine Then
                    ils.Delete
                End If
            Next ils
        End If
    Next para

    MsgBox "All horizontal lines deleted."
End Sub
