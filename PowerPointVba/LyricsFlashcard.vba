Sub CreatePoemSlides()
    Dim filePath As String
    Dim lineText As String
    Dim slide As slide
    Dim textBox As Shape
    Dim fso As Object, ts As Object

    filePath = "./poem.txt" '  Update this path

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
            .AdvanceTime = 15
        End With
    Next s
End Sub

