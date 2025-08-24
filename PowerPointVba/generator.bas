Option Explicit

Const ppSaveAsOpenXMLMacroEnabled As Long = 25
Const songForPowerpoint As String = "poem.txt"
Const TEXTBOX_LEFT As Single = 50
Const TEXTBOX_TOP As Single = 100
Const TEXTBOX_WIDTH As Single = 600
Const TEXTBOX_HEIGHT As Single = 400
Const FONT_SIZE As Integer = 60
Const FONT_NAME As String = "Calibri"

Sub DeleteAllSlides()
    Dim i As Integer
    With ActivePresentation
        For i = .Slides.Count To 1 Step -1
            .Slides(i).Delete
        Next i
    End With
End Sub

Sub GenerateLyricsPptm(songName As String)
' This automates the creation of a PowerPoint presentation from a plain text file.
' Each line in the text file becomes the content of a new slide in the presentation.
' The resulting presentation is saved as a macro-enabled PowerPoint file (.pptm).
    Dim fso As Object, ts As Object
    Dim lineText As String, lines As Collection
    Dim basePath As String, songPath As String, outputName As String
    Dim newPres As Presentation, slide As slide, shp As Shape
    Dim line As Variant, stem As String

    ' Set up file system and paths
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = ActivePresentation.Path
    songPath = basePath & "\" & songName

    ' Derive stem and output name
    stem = fso.GetBaseName(songName)
    outputName = basePath & "\" & stem & ".pptm"

    ' Read lines from poem.txt
    If Not fso.FileExists(songPath) Then
        MsgBox songName & " not found in " & basePath, vbCritical
        Exit Sub
    End If

    Set ts = fso.OpenTextFile(songPath, 1)
    Set lines = New Collection
    Do While Not ts.AtEndOfStream
        lineText = Trim(ts.ReadLine)
        If Len(lineText) > 0 Then lines.Add lineText
    Loop
    ts.Close

    ' Create new presentation
    Set newPres = Presentations.Add(msoTrue)

    ' Add slides for each line
    For Each line In lines
        Set slide = newPres.Slides.Add(newPres.Slides.Count + 1, ppLayoutBlank)
        Set shp = slide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
            Left:=TEXTBOX_LEFT, Top:=TEXTBOX_TOP, _
            Width:=TEXTBOX_WIDTH, Height:=TEXTBOX_HEIGHT)
    With shp.TextFrame
        .TextRange.Text = line
        .TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .TextRange.Font.Size = FONT_SIZE
        .TextRange.Font.Name = FONT_NAME
        .AutoSize = ppAutoSizeShapeToFitText
    End With
    
        ' Center shape manually
        shp.Left = (newPres.PageSetup.SlideWidth - shp.Width) / 2
        shp.Top = (newPres.PageSetup.SlideHeight - shp.Height) / 2
        ' Set slide timing
        slide.SlideShowTransition.AdvanceOnTime = msoTrue
        slide.SlideShowTransition.AdvanceTime = 5
    Next line

    ' Save as pptm
    newPres.SaveAs outputName, ppSaveAsOpenXMLMacroEnabled
    
    MsgBox "Presentation saved as " & outputName, vbInformation
End Sub

Sub RunLyricsAutomation()
    Call DeleteAllSlides
    Call GenerateLyricsPptm(songForPowerpoint)
End Sub

