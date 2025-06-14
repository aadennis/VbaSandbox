' This module is designed to tidy up a Word document exported from OneNote.
' It sets margins, resizes images, and adds footers with specific formatting.


Attribute VB_Name = "TidyOneNoteExportWithFooter"
Sub TidyOneNoteExportWithFooter()
    Const MAX_WIDTH_CM As Single = 16
    Const MAX_HEIGHT_CM As Single = 20
    Const MARGIN_CM As Single = 1.0

    Dim pic As InlineShape
    Dim sec As Section
    Dim footerFirst As HeaderFooter
    Dim footerRest As HeaderFooter
    Dim linePara As Paragraph
    Dim textPara As Paragraph
    Dim r As Range

    ' Set margins for all sections
    For Each sec In ActiveDocument.Sections
        With sec.PageSetup
            .TopMargin = CentimetersToPoints(MARGIN_CM)
            .BottomMargin = CentimetersToPoints(MARGIN_CM)
            .LeftMargin = CentimetersToPoints(MARGIN_CM)
            .RightMargin = CentimetersToPoints(MARGIN_CM)
            .DifferentFirstPageHeaderFooter = True
        End With
    Next sec

    ' Resize and center images
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

    ' Add footers
    For Each sec In ActiveDocument.Sections
        ' === First page footer: full file path, left aligned ===
        Set footerFirst = sec.Footers(wdHeaderFooterFirstPage)
        With footerFirst.Range
            .Text = ""

            ' Divider line (bottom border of an empty paragraph)
            Set linePara = .Paragraphs.Add
            With linePara.Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorGray25
            End With
            linePara.Range.Text = ""
            linePara.Range.Font.Size = 1
            linePara.Range.ParagraphFormat.SpaceAfter = 3

            ' Footer text paragraph
            Set textPara = .Paragraphs.Add
            With textPara.Range
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
                .Font.Name = "Calibri"
                .Font.Size = 8
                .Fields.Add Range:=textPara.Range, Type:=wdFieldFileName, Text:="\p", PreserveFormatting:=False
            End With
        End With

        ' === Rest of pages: page x of y, right aligned ===
        Set footerRest = sec.Footers(wdHeaderFooterPrimary)
        With footerRest.Range
            .Text = ""

            ' Divider line
            Set linePara = .Paragraphs.Add
            With linePara.Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorGray25
            End With
            linePara.Range.Text = ""
            linePara.Range.Font.Size = 1
            linePara.Range.ParagraphFormat.SpaceAfter = 3

            ' Page number text
            Set textPara = .Paragraphs.Add
            With textPara.Range
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Name = "Calibri"
                .Font.Size = 8
                Set r = textPara.Range.Duplicate
                r.Collapse Direction:=wdCollapseStart
                r.Fields.Add Range:=r, Type:=wdFieldPage
                r.InsertAfter " of "
                r.Collapse Direction:=wdCollapseEnd
                r.Fields.Add Range:=r, Type:=wdFieldNumPages
            End With
        End With
    Next sec

    MsgBox "Document tidied: margins set, images resized, footer updated.", vbInformation
End Sub

