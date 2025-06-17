Sub DeleteFooter()
' Delete the footer from the first section of the active document
   ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Delete
End Sub