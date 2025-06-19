' This module contains macros to format text from ChatGPT responses in Word.
Sub FormatChatGPTText()
    Call RemoveTextBeforeYouSaid   ' Rule 1
    Call ApplyChatStyles           ' Rule 2
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

Sub ApplyChatStyles()
    ' Apply specific styles to paragraphs based on the speaker in a ChatGPT conversation.
    ' It assumes the text "You said:" indicates the user's input and "ChatGPT said:" indicates the AI's response.
    ' Right now, the styles are placeholders

    Dim para As Paragraph
    Dim currentSpeaker As String
    currentSpeaker = ""

    For Each para In ActiveDocument.Paragraphs
        Dim txt As String
        txt = Trim(para.Range.Text)
        
        If txt = "You said:" Then
            currentSpeaker = "User"
        ElseIf txt = "ChatGPT said:" Then
            currentSpeaker = "GPT"
        Else
            ' Only apply style if we're inside a speaker block
            Select Case currentSpeaker
                Case "User"
                    On Error Resume Next
                    para.Range.Style = "UserChat"
                    On Error GoTo 0
                Case "GPT"
                    On Error Resume Next
                    para.Range.Style = "GPTChat"
                    On Error GoTo 0
            End Select
        End If
    Next para
End Sub

