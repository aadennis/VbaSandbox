' README:
' VsCode is the repository for this code.
' The .pptm file is just a vehicle for presentation - it should always be possible
' to create a working .pptm from 1. this vba code and 2. a source file containing
' the lyrics of a given song.
' Also note that although VsCode does syntax colouring, the code is not executable
' from VsCode.

' === CONFIGURATION BLOCK ===
Const FILE_NAME As String = "sample_lyrics.txt"
Const TEXTBOX_LEFT As Single = 50
Const TEXTBOX_TOP As Single = 100
Const TEXTBOX_WIDTH As Single = 600
Const TEXTBOX_HEIGHT As Single = 400
Const FONT_SIZE As Integer = 60


Sub DeleteAllSlides()
    Dim i As Integer
    With ActivePresentation
        For i = .Slides.Count To 1 Step -1
            .Slides(i).Delete
        Next i
    End With
End Sub

Sub CreateSlidesForLyrics()
    DeleteAllSlides ' Clear existing slides before generating new ones

    Dim filePath As String
    filePath = ActivePresentation.Path & "\" & FILE_NAME


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
            Left:=TEXTBOX_LEFT, Top:=TEXTBOX_TOP, _
            Width:=TEXTBOX_WIDTH, Height:=TEXTBOX_HEIGHT)


        With textBox.TextFrame.TextRange
            .Text = lineText
            .Font.Size = FONT_SIZE
            .ParagraphFormat.ALIGNMENT = ppAlignLeft
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



