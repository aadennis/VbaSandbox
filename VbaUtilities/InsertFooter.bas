Sub InsertFooter()
    ' Insert a footer with the document path in lowercase, left-aligned.
    ' That path is the full path of the document, starting from '/Documents/',
    ' and excludes the root directory.

    ' Delete any existing footer first
    If Len(ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text) > 1 Then
        ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Delete
        MsgBox "Existing footer deleted.", vbInformation
    End If

    Dim filePathField As Field
    Dim footerRange As Range
    Dim mySection As Section
    Set mySection = ActiveDocument.Sections(1)

    ' Get the footer range
    Set footerRange = mySection.Footers(wdHeaderFooterPrimary).Range
    footerRange.Font.Name = "Arial"
    footerRange.Font.Size = 9
    footerRange.Text = ""

    ' Insert full document path in lowercase, left-aligned
    footerRange.Collapse wdCollapseStart

    Set filePathField = footerRange.Fields.Add(footerRange, wdFieldFileName, "\p")
    filePathField.Update
    filePathText = filePathField.Result

    ' Create RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "/Documents/.*"  ' Match everything from '/Documents/' onward
    regex.IgnoreCase = True
    regex.Global = True

    ' Execute regex and extract match
    If regex.Test(filePathText) Then
        Set match = regex.Execute(filePathText)
        filePathText = match(0)  ' Assign matched text to variable
    Else
        MsgBox "Error: '/Documents/' not found in the path.", vbExclamation
    End If

    ' Remove the field after extracting the path
    filePathField.Delete

    footerRange.Text = LCase(filePathText) ' Convert to lowercase
    footerRange.ParagraphFormat.Alignment = wdAlignParagraphLeft

    ' Update fields
    footerRange.Fields.Update

    InsertHeaderPageNumber()
End Sub

