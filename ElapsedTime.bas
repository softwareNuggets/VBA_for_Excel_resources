Sub main3()
    'written by:  Software Nuggets
    'Youtube: https://www.youtube.com/shorts/a0WaID20VDU
    'Github: https://github.com/softwareNuggets/VBA_for_Excel_resources
    
    'look for the file:    ElapsedTime.bas

    
    'Sheet3 and Data Sheets must exist for this code to work
    
    Dim ws As Worksheet
    Set ws = Worksheets("Sheet3")
    
    Dim ws2 As Worksheet
    Set ws2 = Worksheets("Data")
    
    Dim r As Integer
    Dim i As Integer
    
    Dim startTime As Single
    Dim endTime As Single
    Dim elapsedSeconds As Single
    Dim elapsedMinutes As Integer
    Dim elapsedSecondsInMinutes As Integer
    Dim elapsedMilliseconds As Integer
    
    startTime = Timer
    
    r = 3
    
    While ws.Cells(r, 4) <> ""
    
        n = Replace(ws.Cells(r, 4), " **", "")
        ws.Cells(r, 10) = n
        
        i = 1
        While ws2.Cells(i, 4) <> ""
            d = ws2.Cells(i, 4)
            If (n = d) Then
                ws.Cells(r, 10) = n
            End If
        
            i = i + 1
        Wend
        
        r = r + 1
    Wend
    
    endTime = Timer
    
    'calculate elapsed time in seconds
    elapsedSeconds = endTime - startTime
    
    'calculate elapsed time in minutes
    elapsedMinutes = Int(elapsedSeconds / 60)
    
    
    'calculate elapsed time in seconds
    'within the minute
    elapsedSecondsInMinutes = Int(elapsedSeconds) Mod 60
    
    'calculate elapsed time in milliseconds
    elapsedMilliseconds = (elapsedSeconds - Int(elapsedSeconds)) * 1000
    
    selapsed = selapsed & elapsedMilliseconds & " milliseconds"
    
    MsgBox selapsed
    
End Sub
