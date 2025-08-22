Option Explicit
Const ppSaveAsOpenXMLMacroEnabled = 25

Sub CreatePresentationFromText(filePath As String)
    Dim newPres As Presentation
    Dim slideIndex As Integer
    Dim lineText As String
    Dim fileNum As Integer

    ' Create new presentation
    Set newPres = Presentations.Add
    slideIndex = 1

    ' Open text file
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Read each line and create a slide
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        With newPres.Slides.Add(slideIndex, ppLayoutText)
            .Shapes(1).TextFrame.TextRange.Text = lineText
        End With
        slideIndex = slideIndex + 1
    Loop

    Close #fileNum

    ' Save as .pptm next to input file
    Dim outputPath As String
    outputPath = Left(filePath, InStrRev(filePath, ".")) & "pptm"
    newPres.SaveAs outputPath, ppSaveAsOpenXMLMacroEnabled
    newPres.Close
End Sub

