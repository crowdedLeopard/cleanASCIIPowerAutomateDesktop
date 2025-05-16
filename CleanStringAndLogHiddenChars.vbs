' *****************************************************************************
' * WARNING: This script is a Proof of Concept (POC) and provided as an    *
' * example only. It may not be suitable for production use without       *
' * further testing and modification. Use at your own risk.               *
' *****************************************************************************

' Script to clean strings by removing hidden/non-printable characters
' Originally whipped up: May 2025
' Last tweaked: May 16, 2025
' By: Chris Mathias

'-------------------------------------------------------------------
' Here's a function that takes a string, strips out any weird
' hidden characters, and also keeps a log of what it found.
' It gives back an array: first item is the clean string,
' second is the log of those pesky hidden characters.
'-------------------------------------------------------------------
Function CleanStringAndLogHiddenChars(someText)
    ' Let's get our ducks in a row - a few variables we'll need
    Dim cleanedUpText, reportOfHiddenChars
    Dim i, charAsciiValue
    Dim currentCharacter
    
    ' Start with a clean slate
    cleanedUpText = ""
    reportOfHiddenChars = ""
    
    ' Alright, let's go through the input text, character by character
    For i = 1 To Len(someText)
        ' Grab the current character and figure out its ASCII value
        currentCharacter = Mid(someText, i, 1)
        charAsciiValue = Asc(currentCharacter)
        
        ' Is this a regular, printable character?
        ' (We're talking ASCII codes from space ' ' to tilde '~')
        If charAsciiValue >= 32 And charAsciiValue <= 126 Then
            ' Looks good! Add it to our clean string.
            cleanedUpText = cleanedUpText & currentCharacter
        Else
            ' Uh oh, found a hidden one. Let's make a note of it.
            reportOfHiddenChars = reportOfHiddenChars & "Found at position " & i & ": ASCII value " & charAsciiValue & vbCrLf
        End If
    Next
    
    ' All done! Package up the clean text and our report.
    CleanStringAndLogHiddenChars = Array(cleanedUpText, reportOfHiddenChars)
End Function

'-------------------------------------------------------------------
' Let's see this thing in action...
'-------------------------------------------------------------------
Dim myTestString
Dim processingOutcome
Dim niceAndCleanString
Dim logOfWhatWasRemoved

' Here's some sample text with a few sneaky characters thrown in
myTestString = "This is a test string with some hidden bits: " & Chr(0) & " and " & Chr(7) & " and also " & Chr(31) & ". All done."

' Let our function do its magic
processingOutcome = CleanStringAndLogHiddenChars(myTestString)

' Now, let's unpack the results
niceAndCleanString = processingOutcome(0)
logOfWhatWasRemoved = processingOutcome(1)

' Show the user what we got
MsgBox "Here's the cleaned up string:" & vbCrLf & niceAndCleanString
MsgBox "And here's the log of removed characters:" & vbCrLf & logOfWhatWasRemoved
