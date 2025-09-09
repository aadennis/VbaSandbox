Sub ConvertSelectedListToTwoColumnTable()
    ' This is not working - TBD
    Dim selRange As Range
    Dim para As Paragraph
    Dim entries() As String
    Dim entryCount As Long
    Dim i As Long, rows As Long
    Dim tbl As Table

    Set selRange = Selection.Range

    ' Step 1: Collect non-empty paragraphs from selection

       
    Dim lines() As String
    lines = Split(Selection.Range.Text, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            entryCount = entryCount + 1
            ReDim Preserve entries(1 To entryCount)
            entries(entryCount) = Trim(lines(i))
        End If
    Next i

    If entryCount = 0 Then
        MsgBox "No valid entries found in selection.", vbExclamation
        Exit Sub
    End If

    ' Step 2: Calculate number of rows for 2 columns
    rows = Int((entryCount + 1) / 2)

    ' Step 3: Replace selection with a table
    selRange.Text = ""
    Set tbl = selRange.Tables.Add(Range:=selRange, NumRows:=rows, NumColumns:=2)
    tbl.Borders.Enable = True

    ' Step 4: Fill table cells
    For i = 1 To entryCount
        If i <= rows Then
            tbl.cell(i, 1).Range.Text = entries(i)
        Else
            tbl.cell(i - rows, 2).Range.Text = entries(i)
        End If
    Next i
End Sub
