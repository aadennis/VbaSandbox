Attribute VB_Name = "WordCountTests"
Option Explicit
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub PrintTestResults(testName As String, expectedCount As Integer, actualCount As Integer)
    
    Debug.Print "[" & testName & "][" & Now & "]"
    
    If (expectedCount <> actualCount) Then
        Debug.Print "!!!!!! Expected:[" & expectedCount & "]; Actual:[" & actualCount & "] !!!!!"
    Else
        Debug.Print "OK."
    End If
    Debug.Print "......................"
    Debug.Print
    
End Sub
Sub PrintTestResults2(testName As String, expectedResult As Boolean, actualResult As Boolean)
    
    Debug.Print "[" & testName & "][" & Now & "]"
    
    If (expectedResult <> actualResult) Then
        Debug.Print "!!!!!! Expected:[" & expectedResult & "]; Actual:[" & actualResult & "] !!!!!"
    Else
        Debug.Print "OK."
    End If
    Debug.Print "......................"
    Debug.Print
    
End Sub


Sub TestCountOfTimesOneWordOccursInALine()
    Dim WordToFind As String
    Dim Line As String
    Dim delimiter As String
    Dim expectedCount As Integer
    Dim actualCount As Integer
    
    ' arrange...
    delimiter = ";"
    WordToFind = "Curiosity"
    Line = "Beagle;Rover"
    ' act...
    actualCount = CountOfTimesOneWordOccursInALine(WordToFind, Line, delimiter)
    ' assert...
    PrintTestResults "Test 1", 0, actualCount
    
    delimiter = ";"
    WordToFind = "Curiosity"
    Line = "Curiosity;Rover"
    actualCount = CountOfTimesOneWordOccursInALine(WordToFind, Line, delimiter)
    PrintTestResults "Test 2", 1, actualCount
    
    delimiter = ":"
    WordToFind = "Curiosity"
    Line = "Curiosity:Curiosity"
    actualCount = CountOfTimesOneWordOccursInALine(WordToFind, Line, delimiter)
    PrintTestResults "Test 3", 2, actualCount
    
    delimiter = ";"
    WordToFind = "Curiosity"
    Line = "Curiosity:Curiosity"
    actualCount = CountOfTimesOneWordOccursInALine(WordToFind, Line, delimiter)
    PrintTestResults "Test 4", 0, actualCount
    
    delimiter = ";"
    WordToFind = "Curiosity"
        Line = "Curiosity:Curiosity;Beagle"
    actualCount = CountOfTimesOneWordOccursInALine(WordToFind, Line, delimiter)
    PrintTestResults "Test 5", 66, actualCount
    
    delimiter = " "
    WordToFind = "Beagle"
    Line = "Curiosity Curiosity Beagle"
    actualCount = CountOfTimesOneWordOccursInALine(WordToFind, Line, delimiter)
    PrintTestResults "Test 6", 66, actualCount
    
    'Clear the immediate window...
    'SendKeys "^g ^a {DEL}"
    
End Sub

Sub TestTokenDelimiterTokenFoundInLine()
    Dim WordToFind As String
    Dim Line As String
    Dim delimiter As String
    Dim actualResult As Boolean
    
    
    ' arrange...
    delimiter = ":"
    Line = "Coffee Tea FishAndChips : FishAndChips Pudding More Tea"
    ' act...
    actualResult = TokenDelimiterTokenFoundInLine(Line, delimiter)
    ' assert...
    PrintTestResults2 "Test 1", True, actualResult
    
    delimiter = "$"
    Line = "Coffee Tea FishAndChips : FishAndChips Pudding More Tea"
    actualResult = TokenDelimiterTokenFoundInLine(Line, delimiter)
    PrintTestResults2 "Test 2", False, actualResult
    
    delimiter = ":"
    Line = "Coffee Tea FishAndChipsTwice : FishAndChips Pudding More Tea"
    actualResult = TokenDelimiterTokenFoundInLine(Line, delimiter)
    PrintTestResults2 "Test 3", False, actualResult
    
    delimiter = ":"
    Line = "Coffee Tea FishAndChipsTwice:FishAndChips Pudding More Tea"
    actualResult = TokenDelimiterTokenFoundInLine(Line, delimiter)
    PrintTestResults2 "Test 4", False, actualResult
    
    delimiter = ":"
    Line = vbNullString
    actualResult = TokenDelimiterTokenFoundInLine(Line, delimiter)
    PrintTestResults2 "Test 5", False, actualResult
    
End Sub
