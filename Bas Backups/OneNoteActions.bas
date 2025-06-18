Attribute VB_Name = "OneNoteActions"
' This macro attaches a custom template to the active document and updates the styles.
' It overcomes the default template settings (normal.dotm) and applies the custom styles defined in the template.
Sub AttachCustomTemplateAndUpdate()
    Dim tmplPath As String
    Dim doc As Document
    Dim userProfile As String

    userProfile = Environ("USERPROFILE")
    Template = "OneNote_Styled_Template.dotm"
    tmplPath = userProfile & "\AppData\Roaming\Microsoft\Templates\WordStandards\" & Template
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

    msg = "Custom template '" & Template & "' has been attached successfully." & vbCrLf & _
          "Styles have been updated to match the template settings."
    MsgBox msg, vbInformation
End Sub

Sub InsertLocalPathAndPageNumbersFormatted()
    ' Insert a footer with the local document path and page numbers formatted.
    ' The folder path is left-deleted, due to OneDrive sync issues*, so show only
    ' the path from "Documents/" onward, and format the footer with a right-aligned
    ' tab stop for the page numbers.
    ' * - OneDrive insists on retaining the web path, not the local path. Better to
    ' limit the path to the local "Documents/" folder, regardless of the parent folder.
    Dim docPath As String
    Dim sec As Section
    Dim footerRange As Range
    Dim tabPos As Single
    Dim regex As Object
    Dim matches As Object

    ' Get local full path
    docPath = ActiveDocument.FullName

    ' Extract from "Documents/" onward using regex
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "Documents/.*"
    regex.IgnoreCase = True
    regex.Global = False

    If regex.Test(docPath) Then
        Set matches = regex.Execute(docPath)
        docPath = matches(0)
    End If

    ' Define a right-aligned tab stop at the right margin
    tabPos = CentimetersToPoints(17) ' Adjust to match your page layout

    For Each sec In ActiveDocument.Sections
        Set footerRange = sec.Footers(wdHeaderFooterPrimary).Range
        With footerRange
            .text = "" ' Clear existing footer content
            .ParagraphFormat.TabStops.ClearAll
            .ParagraphFormat.TabStops.Add Position:=tabPos, Alignment:=wdAlignTabRight

            ' Add a horizontal line above using paragraph border
            .ParagraphFormat.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            .ParagraphFormat.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
            .ParagraphFormat.Borders(wdBorderTop).Color = wdColorGray25

            ' Insert path (left) and tab + page numbering (right)
            .InsertAfter docPath & vbTab & "Page "
            .Fields.Add Range:=.Characters.Last, Type:=wdFieldPage
            .InsertAfter " of "
            .Fields.Add Range:=.Characters.Last, Type:=wdFieldNumPages

            ' Format
            .Font.Name = "Calibri"
            .Font.Size = 9
        End With
    Next sec
End Sub

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

End Sub
    



