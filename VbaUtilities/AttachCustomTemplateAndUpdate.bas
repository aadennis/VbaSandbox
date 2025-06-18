' Attach a custom template to the active document and update the styles.
' It overrides the default template settings (normal.dotm) and applies the custom styles defined in the template.
Sub AttachCustomTemplateAndUpdate()
    Dim tmplPath As String
    Dim doc As Document
    Dim userProfile As String

    userProfile = Environ("USERPROFILE")
    template= "OneNote_Styled_Template.dotm"
    tmplPath = userProfile & "\AppData\Roaming\Microsoft\Templates\WordStandards\" & template
    Set doc = ActiveDocument

    ' Attach the custom template
    doc.AttachedTemplate = tmplPath

    ' Update styles to match template
    doc.UpdateStylesOnOpen = True
    doc.UpdateStyles

    ' Optional: set margins (example = Narrow)
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(1.27)
        .BottomMargin = CentimetersToPoints(1.27)
        .LeftMargin = CentimetersToPoints(1.27)
        .RightMargin = CentimetersToPoints(1.27)
    End With

    msg = "Custom template '" & template & "' has been attached successfully." & vbCrLf & _
          "Styles have been updated to match the template settings." 
    MsgBox msg, vbInformation
End Sub