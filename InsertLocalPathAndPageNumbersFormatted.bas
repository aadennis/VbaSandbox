Sub InsertLocalPathAndPageNumbersFormatted()
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
    regex.Pattern = "Documents/.*"
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
            .Text = "" ' Clear existing footer content
            .ParagraphFormat.TabStops.ClearAll
            .ParagraphFormat.TabStops.Add Position:=tabPos, Alignment:=wdAlignTabRight

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
