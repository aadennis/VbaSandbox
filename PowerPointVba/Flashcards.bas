Option Explicit

Const ppSaveAsOpenXMLMacroEnabled As Long = 25
Const TEXTBOX_LEFT As Single = 50
Const TEXTBOX_TOP As Single = 100
Const TEXTBOX_WIDTH As Single = 600
Const TEXTBOX_HEIGHT As Single = 400
Const FONT_SIZE As Integer = 60
Const FONT_NAME As String = "Calibri"
Const ADVANCE_TIME As Integer = 12


' ********************************************************************************************
' * VBA Module
' *
' * Purpose:
' * The main function of this module is to generate Flashcards as part of a PPTM.
' * Typically, its job is to help remember song lyrics, by starting with a delay of n seconds,
' * then presenting a line - did you remember the line?
' * More formally...
' * This module automates the creation of PowerPoint presentations from plain text files.
' * Each line in the text file becomes the content of a new slide in the presentation.
' * The resulting presentation is saved as a macro-enabled PowerPoint file (.pptm).
' *
' * Workflow Overview:
' * 1. **Delete Existing Slides**:
' *    - The `DeleteAllSlides` subroutine removes all slides from the active presentation.  *
' *                                                                                         *
' * 2. **Generate New Presentation**:
' *    - The `GenerateLyricsPptm` subroutine performs the following steps:
' *      a. Reads the specified text file line by line.                                      *
' *      b. Creates a new PowerPoint presentation.                                           *
' *      c. Adds a new slide for each line of text, inserting the text into a centered       *
' *         text box with predefined dimensions, font, and alignment.                        *
' *      d. Sets slide transition timing to 5 seconds per slide.                             *
' *      e. Saves the presentation as a `.pptm` file in the same directory as the text file. *
' *                                                                                          *
' * 3. **Run Automation**:                                                                  *
' *    - The `RunLyricsAutomation` subroutine orchestrates the process by:                  *
' *      a. Deleting all existing slides.                                                    *
' *      b. Calling `GenerateLyricsPptm` with the predefined text file (`poem.txt`).         *
' *                                                                                          *
' * Key Constants:                                                                           *
' * - `ppSaveAsOpenXMLMacroEnabled`: File format for saving macro-enabled PowerPoint files.  *
' * - `TEXTBOX_*`: Dimensions and positioning for the text box on each slide.                *
' * - `FONT_*`: Font name and size for the text.                                             *
' *                                                                                          *
' * Usage:                                                                                   *
' * - Place the text file (`poem.txt`) in the same directory as the PowerPoint presentation. *
' * - Run the `RunLyricsAutomation` macro to generate the presentation.                      *
' *                                                                                          *
' ********************************************************************************************

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
        lineText = ts.ReadLine
        If IsLyricLine(lineText) Then lines.Add Trim(lineText)
    Loop
    ts.Close

    ' Create new presentation
    Set newPres = Presentations.Add(msoTrue)

    ' Add intro slide with source metadata
    Dim introSlide As slide
    Set introSlide = newPres.Slides.Add(1, ppLayoutText)

    With introSlide.Shapes(1).TextFrame.TextRange
        .Text = songName
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
        .ParagraphFormat.Alignment = ppAlignCenter
    End With

    On Error Resume Next
    introSlide.Shapes(2).Delete
    On Error GoTo 0

    ' Add slides for each line
    Dim i As Integer
    For i = 1 To lines.Count
        ' Add lyric slide
        Set slide = newPres.Slides.Add(newPres.Slides.Count + 1, ppLayoutBlank)
        Set shp = slide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
            Left:=TEXTBOX_LEFT, Top:=TEXTBOX_TOP, _
            Width:=TEXTBOX_WIDTH, Height:=TEXTBOX_HEIGHT)
        With shp.TextFrame
            .TextRange.Text = lines(i)
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
        slide.SlideShowTransition.AdvanceTime = ADVANCE_TIME
    
        ' Add separator slide
        Set slide = newPres.Slides.Add(newPres.Slides.Count + 1, ppLayoutBlank)
        Set shp = slide.Shapes.AddShape(Type:=msoShapeRectangle, _
            Left:=0, Top:=newPres.PageSetup.SlideHeight / 2 - 10, _
            Width:=newPres.PageSetup.SlideWidth, Height:=20)
        With shp
            .Fill.ForeColor.RGB = RGB(0, 0, 0) ' Black line
            .line.Visible = msoFalse
        End With
        slide.SlideShowTransition.AdvanceOnTime = msoTrue
        slide.SlideShowTransition.AdvanceTime = ADVANCE_TIME
    Next i

    ' Save as pptm
    newPres.SaveAs outputName, ppSaveAsOpenXMLMacroEnabled
    
    MsgBox "Presentation saved as " & outputName, vbInformation
End Sub

Sub SetSlideTimings()
    Dim s As slide
    For Each s In ActivePresentation.Slides
        With s.SlideShowTransition
            .AdvanceOnTime = msoTrue
            .AdvanceTime = ADVANCE_TIME
        End With
    Next s
End Sub

Function IsBlankLine(lineText As String) As Boolean
    IsBlankLine = (Len(Trim(lineText)) = 0)
End Function

Function IsChordLine(lineText As String) As Boolean
    Dim knownChords As Variant
    knownChords = Array("A", "B", "C", "D", "E", "F", "G", _
                        "Am", "Bm", "Cm", "Dm", "Em", "Fm", "Gm", _
                        "A7", "B7", "C7", "D7", "E7", "F7", "G7", _
                        "Amaj7", "C#m", "F#m", "G#7", "Dmaj7") ' Expand as needed

    Dim tokens() As String
    Dim token As String
    Dim matchCount As Integer
    Dim i As Integer, j As Integer

    lineText = Trim(lineText)
    If Len(lineText) = 0 Then
        IsChordLine = False
        Exit Function
    End If

    tokens = Split(lineText)
    matchCount = 0

    For i = LBound(tokens) To UBound(tokens)
        token = Replace(tokens(i), "-", "") ' Remove dashes
        token = Replace(token, "–", "")     ' Remove en-dashes
        token = Replace(token, "—", "")     ' Remove em-dashes
        token = Trim(token)

        For j = LBound(knownChords) To UBound(knownChords)
            If StrComp(token, knownChords(j), vbTextCompare) = 0 Then
                matchCount = matchCount + 1
                Exit For
            End If
        Next j
    Next i

    IsChordLine = (matchCount >= 2)
End Function

Function IsLyricLine(lineText As String) As Boolean
    IsLyricLine = (Not IsBlankLine(lineText)) And (Not IsChordLine(lineText))
End Function

Function ReadConfigValue(key As String, configPath As String) As String
    Dim fso As Object, ts As Object, line As String, parts() As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(configPath) Then
        MsgBox "Config file not found: " & configPath, vbCritical
        End
    End If

    Set ts = fso.OpenTextFile(configPath, 1)
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If InStr(line, "=") > 0 Then
            parts = Split(line, "=")
            If UCase(Trim(parts(0))) = UCase(key) Then
                ReadConfigValue = Trim(parts(1))
                ts.Close
                Exit Function
            End If
        End If
    Loop
    ts.Close
    MsgBox "Key '" & key & "' not found in config file.", vbCritical
    End
End Function

Function GetFlashcardSource() As String
    ' Retrieves the flashcard source from the config file.
    Dim configPath As String
    Dim flashcardSource As String

    configPath = ActivePresentation.Path & "\config.txt"
    flashcardSource = ReadConfigValue("FLASHCARD_SOURCE", configPath)
    
    GetFlashcardSource = flashcardSource
End Function

Sub RunLyricsAutomation()
    Dim flashcardSource As String
    flashcardSource = GetFlashcardSource()

    Call DeleteAllSlides
    Call GenerateLyricsPptm(flashcardSource)
    Call SetSlideTimings
End Sub
