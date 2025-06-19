' This module contains macros to format text from ChatGPT responses in Word.
Sub FormatChatGPTText()
    ' Master macro — calls each rule one by one
    Call RemoveTextBeforeYouSaid
    ' Future rules can be added here as more Call statements
End Sub

Sub RemoveTextBeforeYouSaid()
    ' Remove all text before the first occurrence of "You said:" in the active document.
    ' It is useful for cleaning up ChatGPT responses that may have introductory text.
    ' It assumes the text "You said:" is used to indicate the start of the user's
    ' input in the conversation.

    Dim doc As Document
    Dim rng As Range
    Dim searchText As String

    Set doc = ActiveDocument
    searchText = "You said:"
    
    Set rng = doc.Content
    With rng.Find
        .Text = searchText
        .Forward = True
        .MatchCase = False
        .Execute
    End With

    If rng.Find.Found Then
        doc.Range(0, rng.Start).Delete
        Debug.Print "Rule 1 applied: removed text before 'You said:'"
    Else
        Debug.Print "Rule 1 skipped: 'You said:' not found"
    End If
End Sub

