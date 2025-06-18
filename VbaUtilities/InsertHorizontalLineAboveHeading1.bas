Sub InsertHorizontalLineAboveHeading1()
    ' Insert a horizontal line above each Heading 1 style paragraph in the active document.
    ' This uses paragraph borders to create a horizontal line effect.
    ' This helps distinguish between topics in the document imported from Google Docs.
    ' It shouuld be run after executing DeleteAllHorizontalLines() - see comments there.
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "Heading 1" Then
            para.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            para.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
            para.Borders(wdBorderTop).Color = wdColorAutomatic
        End If
    Next para
End Sub