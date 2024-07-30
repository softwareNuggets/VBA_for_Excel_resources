'video for source code
'https://www.youtube.com/watch?v=fPV4CFUN4xU&list=PLRU_t-SgTrYiymghXbBJynOMBAqgdkN4h&index=11

Sub RemoveRange(sheetName As String, rangeName As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim wsExists As Boolean
    Dim namedRangeExists As Boolean
    
    wsExists = False
    namedRangeExists = False

    ' Check if the worksheet exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            wsExists = True
            Exit For
        End If
    Next ws

    ' If the worksheet does not exist, exit the subroutine
    If Not wsExists Then
        MsgBox "Worksheet '" & sheetName & "' does not exist.", vbExclamation
        Exit Sub
    End If

    ' Check if the named range exists
    On Error Resume Next
    Set rng = ws.Range(rangeName)
    If Not rng Is Nothing Then
        namedRangeExists = True
    End If
    On Error GoTo 0

    ' If the named range does not exist, exit the subroutine
    If Not namedRangeExists Then
        MsgBox "Range '" & rangeName & "' does not exist in worksheet '" & sheetName & "'.", vbExclamation
        Exit Sub
    End If

    ' Delete the range
    rng.Delete
    MsgBox "Range '" & rangeName & "' has been removed from worksheet '" & sheetName & "'.", vbInformation
End Sub
