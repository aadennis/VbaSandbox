'Option Explicit
' This macro removes numbered lines from a Word document.
' It uses a regular expression to identify lines that start with digits followed by a dot and optional whitespace.
' The macro then replaces the matched portion with an empty string, effectively removing the numbering.
' The macro also counts the number of lines modified and displays a message box with the count.

Sub RemoveNumberedLines()
    Dim para As Paragraph
    Dim r As Range
    Dim lineText As String
    Dim regex As Object
    Dim pattern As String
    Dim modifiedText As String
    Dim changedCount As Integer
    changedCount = 0
    
    ' Create regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Regular expression pattern to match digits followed by a dot and optional whitespace
    pattern = "^\d+\.\s*"
    
    regex.IgnoreCase = True
    regex.Global = True
    regex.pattern = pattern
    
    ' Loop through each paragraph in the document
    For Each para In ActiveDocument.Paragraphs
        Set r = para.Range
        lineText = r.Text
        
        ' Check if the line matches the pattern (digit(s) + dot + whitespace)
        If regex.Test(lineText) Then
            ' Replace the matched portion (digits + dot + any whitespace after it) with an empty string
            modifiedText = regex.Replace(lineText, "")
            r.Text = modifiedText
            changedCount = changedCount + 1
        End If
    Next para

    MsgBox changedCount & " line(s) modified.", vbInformation, "Done"
End Sub

' This macro styles what passes for code blocks in a Word document.
' It checks each paragraph for specific monospaced fonts (Courier New, Consolas, or Cascadia Code).
' If a paragraph uses one of these fonts, it sets the font size to 9pt and the background color to black.
' It also adds a border around the paragraph.
' The macro counts the number of paragraphs modified and displays a message box with the count.
' There is one version for the whole document and another for the selection.


Sub StyleCodeBlocksInDocument()
    Dim r As Range
    Dim changedCount As Integer
    Dim para As Paragraph
    changedCount = 0

    For Each para In ActiveDocument.Paragraphs
        Set r = para.Range
        ' Check for monospaced fonts
        If r.Font.Name = "Courier New" Or r.Font.Name = "Consolas" Or r.Font.Name = "Cascadia Code" Then
            r.Font.Size = 9
            r.Shading.BackgroundPatternColor = wdColorBlack
            r.Borders.Enable = True
            r.Borders.OutsideLineStyle = wdLineStyleSingle
            r.Borders.OutsideColor = wdColorOrange
            changedCount = changedCount + 1
        End If
    Next para

    MsgBox changedCount & " code block(s) styled.", vbInformation, "Done"
End Sub

Sub StyleCodeBlocksInSelection()
    Dim r As Range
    Dim changedCount As Integer
    Dim para As Paragraph
    changedCount = 0

    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select some text first.", vbExclamation, "No Selection"
        Exit Sub
    End If

    For Each para In Selection.Paragraphs
        Set r = para.Range
        ' Check for monospaced fonts
        If r.Font.Name = "Courier New" Or r.Font.Name = "Consolas" Or r.Font.Name = "Cascadia Code" Then
            r.Font.Size = 9
            r.Shading.BackgroundPatternColor = wdColorBlack
            r.Borders.Enable = True
            r.Borders.OutsideLineStyle = wdLineStyleSingle
            r.Borders.OutsideColor = wdColorBlack
            changedCount = changedCount + 1
        End If
    Next para

    MsgBox changedCount & " code block(s) styled in selection.", vbInformation, "Done"
End Sub

