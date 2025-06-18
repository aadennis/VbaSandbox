Sub InsertHeaderPageNumber()
    ' Insert a page number in the header, top right (right-aligned).

    ' Delete any existing header content first
    If Len(ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text) > 1 Then
        ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Delete
        MsgBox "Existing header deleted.", vbInformation
    End If

    Dim headerRange As Range
    Set headerRange = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range
    headerRange.Font.Name = "Arial"
    headerRange.Font.Size = 8
    headerRange.Text = ""

    ' Right-align and insert "Page { PAGE }"
    headerRange.ParagraphFormat.Alignment = wdAlignParagraphRight
    headerRange.InsertAfter "Page "
    headerRange.Collapse wdCollapseEnd
    headerRange.Fields.Add headerRange, wdFieldPage
    headerRange.Collapse wdCollapseEnd

    headerRange.Fields.Update
End Sub