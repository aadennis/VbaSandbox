Sub InsertFooter()
    ' Insert a footer with the full document path in lowercase, left-aligned
    Dim filePathField As Field 
    Dim footerRange As Range
    Dim mySection As Section
    Set mySection = ActiveDocument.Sections(1)

    ' Get the footer range
    Set footerRange = mySection.Footers(wdHeaderFooterPrimary).Range
    footerRange.Font.Name = "Arial"
    footerRange.Font.Size = 9
    footerRange.text = ""

    ' Insert full document path in lowercase, left-aligned
    footerRange.Collapse wdCollapseStart
    
    Set filePathField = footerRange.Fields.Add(footerRange, wdFieldFileName, "\p")
    filePathField.Update
    filePathText = filePathField.Result
    
     ' Create RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "/Documents/.*"  ' Match everything from '/Documents/' onward
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
    
    footerRange.text = LCase(filePathText) ' Convert to lowercase
    footerRange.ParagraphFormat.Alignment = wdAlignParagraphLeft

    ' Add a line break before inserting the page number fields
    
End Sub

