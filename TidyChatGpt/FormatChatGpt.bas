' This module contains macros to format text from ChatGPT responses in Word.
Sub FormatChatGPTText()
    Call RemovePreChatText   ' Rule 1
    Call ApplyChatStyles     ' Rule 2
End Sub

Sub RemovePreChatText()
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
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        If txt = "You said:" Then
            currentSpeaker = "User"
        ElseIf txt = "ChatGPT said:" Then
            currentSpeaker = "GPT"
        Else
            ' Only apply style if we're inside a speaker block
            Select Case currentSpeaker
                Case "User"
                    On Error Resume Next
                    para.Range.Style = "userChat"
                    On Error GoTo 0
                Case "GPT"
                    On Error Resume Next
                    para.Range.Style = "GPTChat"
                    
                    On Error GoTo 0
            End Select
        End If
    Next para
End Sub
Sub NewChatGPTStyledDoc()

    ' Create a new document based on a predefined template for ChatGPT style rules.
    ' Word normally expects a template file to be in the user's templates directory,
    ' and not a folder dictated by the user. This is enforced by the Word UI, in that
    ' other folders are not shown in the "New Document" dialog.
    ' However, after this has run, and has created a draft new document based on that
    ' template, you are then in the UI, and can save the document wherever you like.
    Dim templatePath As String
    templatePath = "C:\Users\Dennis\AppData\Roaming\Microsoft\Templates\WordStandards\ChatGPTStyleRules.dotm"
    
    Documents.Add Template:=templatePath, NewTemplate:=False
End Sub

Sub PasteWithLineBreaks()
    ' Paste clipboard content into the active document, ensuring line breaks are preserved.
    ' Context: This is useful when pasting text from ChatGPT or other sources where line breaks
    ' may not be preserved correctly. The macro attempts to paste as plain text and then
    ' fixes line breaks by replacing single line feeds with carriage return + line feed pairs.
    Dim textRange As Range

    ' Try pasting as plain text
    On Error GoTo fallback
    Selection.PasteSpecial DataType:=wdPasteText
    GoTo fixBreaks

fallback:
    MsgBox "Clipboard paste failed — use Notepad workaround instead.", vbExclamation
    Exit Sub

fixBreaks:
    Set textRange = Selection.Range
    textRange.text = Replace(textRange.text, vbLf, vbCrLf)
End Sub