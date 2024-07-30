Sub FillRangeWithRowCol()

    Dim rng As Range
    Dim r As Long, c As Long
    Dim maxRow As Long, maxCol As Long

    Dim ws As Worksheet
    Set ws = Worksheets("Sheet1")
    
    ' Define the range within the worksheet
    Set rng = ws.Range("range_1")

    ' Get the maximum row and
    ' column numbers in the range
    maxRow = rng.Rows.Count
    maxCol = rng.Columns.Count

    ' Loop through each cell in the range
    For r = 1 To maxRow
     For c = 1 To maxCol
        
       ' Set the cell value to "row#_col#"
       rng.Cells(r, c).Value = _
                "r" & CStr(r) & "_c" & CStr(c)
                    
     Next c
    Next r
End Sub
