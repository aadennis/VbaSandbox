' README:
' VsCode is the repository for this code.
' The .pptm file is just a vehicle for presentation - it should always be possible
' to create a working .pptm from 1. this vba code and 2. a source file containing
' the lyrics of a given song.
' Also note that although VsCode does syntax colouring, the code is not executable
' from VsCode.

Sub CreateSlidesForLyrics()
    Dim fileName As String
    Dim filePath As String
    
    fileName = "sample_lyrics.txt" ' FILENAME CONSTANT - update this
    filePath = ActivePresentation.Path & "\" & fileName

    Dim lineText As String
    Dim slide As slide
    Dim textBox As Shape
    Dim fso As Object, ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)

    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        Set slide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
        Set textBox = slide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                               Left:=50, Top:=100, Width:=600, Height:=400)
        With textBox.TextFrame.TextRange
            .Text = lineText
            .Font.Size = 60
            .ParagraphFormat.Alignment = ppAlignLeft
        End With
        slide.FollowMasterBackground = msoFalse
    Loop

    ts.Close
End Sub

Sub SetSlideTimings()
    Dim s As slide
    For Each s In ActivePresentation.Slides
        With s.SlideShowTransition
            .AdvanceOnTime = msoTrue
            .AdvanceTime = 7
        End With
    Next s
End Sub

