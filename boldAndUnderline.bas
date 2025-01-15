Sub FormatMessages()
    Dim l As Long
    Dim r As Long
    
    l = 1
    For r = 1 To 3
        Call BoldAndUnderlinePatternInCell(r, l)
    Next r
End Sub

Sub BoldAndUnderlinePatternInCell(row As Long, col As Long)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim inputCell As Range
    Set inputCell = ws.Cells(row, col)
    
    Dim inputStr As String
    inputStr = inputCell.Value

  ' First, remove any existing formatting
    inputCell.Font.Bold = False
    inputCell.Font.Underline = xlNone
    
    Dim patternStart As Long
    patternStart = InStr(1, inputStr, "~[")
    
    While patternStart > 0
        Dim patternEnd As Long
        patternEnd = InStr(patternStart + 1, inputStr, "]~")
        
        If patternEnd > 0 Then
            ' Found a pattern
            inputCell.Characters(patternStart, patternEnd - patternStart + 2).Font.Bold = True
            inputCell.Characters(patternStart, patternEnd - patternStart + 2).Font.Underline = True
            
            ' Move past the pattern and look for the next one
            patternStart = InStr(patternEnd + 1, inputStr, "~[")
        Else
            ' Pattern end not found, stop looking for more patterns
            patternStart = 0
        End If
    Wend
End Sub


'source code is avialable at ::https://github.com/softwareNuggets/VBA_for_Excel_resources
