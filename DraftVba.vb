Option Explicit

Function CountOfTimesOneWordOccursInALine(WordToFind As String, Line As String, delimiter As String) _
As Integer
' Count the number of times a single word (aka a shorter string) occurs in a line
' (aka a longer string)

    'First do the easy thing - if the search string does not occur in the target string,
    'return count of 0
    If (InStr(Line, WordToFind) = 0) Then
        CountOfTimesOneWordOccursInALine = 0
        Exit Function
    End If
    
    Dim searchStringCount As Integer
    Dim uniqueSet As New Scripting.Dictionary
    Dim wordsInLine() As String
    
    'Got here? Found the search string at least once, so insert into the (unique key) dictionary...
    searchStringCount = 0
    uniqueSet.Add WordToFind, Null
    
    'get the target words into an array so we can walk them...
    wordsInLine = Split(Line, delimiter)
    
    Dim wordInLine As Variant
    For Each wordInLine In wordsInLine
        If uniqueSet.Exists(wordInLine) Then
            searchStringCount = searchStringCount + 1
        End If
    Next wordInLine
    
    CountOfTimesOneWordOccursInALine = searchStringCount
End Function

Function TokenDelimiterTokenFoundInLine(Line As String, delimiter As String) _
As Boolean
' As an example, if I have an example line with this content:
' [Coffee Tea FishAndChips : FishAndChips Pudding More Tea], this returns True because there is a token (FishAndChips)
' immediately followed by the passed delimiter (once any white space is ignored), followed the a second instance of the
' same token value.
' Else (that pattern is not found at least once in the line), return False
' Let's restate the rule (and assume that ":" is the passed delimiter)
' given consecutive tokens in a line, look for the token ":" (assumption that there will be only once such token/delimiter
' in a line. Now, find and save the preceding token, skipping any whitespace. (if none found return)
' Now find the token that immediately follows the ":" (again skipping whitespace, and if none found, again return).
' If the pattern token/delimiter/token occurs in the line, return true, else return false

    'First do the easy thing - if the requested delimiter does not occur in the target string,
    'return False (delimter not found)
    If (InStr(Line, delimiter) = 0) Then
        TokenDelimiterTokenFoundInLine = False
        Debug.Print "Did not find the delimter, returning false"
        
        Exit Function
    End If

    'delimiter was found, now get the tokens before and after
    Dim tokenSet() As String
    Dim delimIndex  As Integer
    
    
    tokenSet = Split(Line)
    Debug.Print Line

    Dim x As Variant
    For Each x In tokenSet
        Debug.Print delimIndex & "[" & x & "]"
        If x = delimiter Then
            Debug.Print "got the delimiter at index " & delimIndex
            Debug.Print "Value is:" & tokenSet(delimIndex)
            'Now test the preceding and next tokens. If they are the same then return true
            If tokenSet(delimIndex - 1) = tokenSet(delimIndex + 1) Then
                TokenDelimiterTokenFoundInLine = True
                Debug.Print "Found the token and pre and post"
            Else
                TokenDelimiterTokenFoundInLine = False
                Debug.Print "Found the delim, but pre and post did not match"
            End If
            Exit Function
            
        End If
        
        delimIndex = delimIndex + 1
    Next
    
    
    
    x = True
    
    





End Function

