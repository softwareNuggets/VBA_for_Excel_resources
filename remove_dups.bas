Attribute VB_Name = "Module1"
Sub remove_dups()

    Dim ws As Worksheet
    Set ws = ActiveSheet        ' Set the worksheet

    'Tools->References->Microsoft Scripting Runtime
    Dim dataDict As Object     ' Create a dictionary to store key-value pairs
    Set dataDict = CreateObject("Scripting.Dictionary")
    
    Dim startRow As Integer
    startRow = 2
    
    Dim use_this_column_index As Integer
    use_this_column_index = 2

    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim cell As Range
    Dim keys As Variant
    Dim key As Variant
    Dim uniqueFieldData As String
    
    ' Find the last row in column C
    lastRow = ws.Cells(ws.Rows.Count, use_this_column_index).End(xlUp).Row
    
    
    ' First pass: Collect all UNIQUE emails
    For i = startRow To lastRow
        Set cell = ws.Cells(i, use_this_column_index)
        
        If Not IsEmpty(cell) Then
            keys = Split(cell.Value, ";")
            
            For Each key In keys
                key = Trim(key)
                
                If key <> "" Then
                    If Not dataDict.Exists(key) Then
                        dataDict.Add key, i
                    End If
                End If
                
            Next key
        End If
    Next i
    
    
    
    
    ' Second pass: Remove duplicates
    For i = lastRow To startRow Step -1
    
        Set cell = ws.Cells(i, use_this_column_index)
        
        If Not IsEmpty(cell) Then
        
            keys = Split(cell.Value, ";")
            uniqueFieldData = ""
            
            For Each key In keys
                key = Trim(key)
                
                If key <> "" Then
                
                    If dataDict(key) = i Then
                    
                        If uniqueFieldData <> "" Then
                            uniqueFieldData = uniqueFieldData & ";"
                        End If
                        
                        uniqueFieldData = uniqueFieldData & key
                    End If
                End If
                
            Next key
            cell.Value = uniqueFieldData
        End If
    Next i
    
    ' Clean up
    dataDict.RemoveAll
    Set dataDict = Nothing
End Sub

