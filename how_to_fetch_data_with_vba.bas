Sub main()

    '->Tools->References->
    'Microsoft ActiveX Data Object  2.8 Library
    
    'GITHUB Source code Is here:
    'https://github.com/softwareNuggets/VBA_for_Excel_resources/blob/main
    '                                  /how_to_fetch_data_with_vba.bas

    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim sql_query As String
    Dim row_counter As Long
    Dim field_counter As Integer
    Dim ws As Worksheet
    
    conn.ConnectionString = Build_SQL_ConnectionString()
    conn.Open
    
    On Error GoTo CloseConnection


    If conn.State <> adStateOpen Then
        Debug.Print "invalid connection string"
        Debug.Print conn.ConnectionString
        End
    End If
    
    
    sql_query = "SELECT product, orderdate, quantity FROM sales"
    
    rs.ActiveConnection = conn
    rs.CursorType = adOpenForwardOnly
    rs.Open sql_query
    
    Set ws = Worksheets("Sheet1")
    
    With ws
        ' clear previous data
        .Cells.ClearContents
        
        ' write header row
        For field_counter = 0 To rs.Fields.Count - 1
            .Cells(1, field_counter + 1).Value = rs.Fields(field_counter).Name
        Next field_counter
        
        ' write data rows
        row_counter = 2
        While Not rs.EOF
            .Cells(row_counter, 1) = rs.Fields("product").Value
            .Cells(row_counter, 2) = rs("orderdate").Value
            .Cells(row_counter, 3) = rs("quantity").Value
            
            rs.MoveNext
            row_counter = row_counter + 1
        Wend
    End With
    
    
CloseConnection:
    If rs.State = adStateOpen Then
        rs.Close
    End If
    conn.Close
    Set conn = Nothing
    Set rs = Nothing
End Sub



Function Build_SQL_ConnectionString() As String

'->Tools->References->
    'Microsoft ActiveX Data Object  2.8 Library
    
    'https://youtube.com/shorts/962bzAIOWt8
    'Excel VBA: How do I connect to SQL Server using VBA? #shorts
    'https://github.com/softwareNuggets/VBA_for_Excel_resources/blob/main
    '                      /how_to_connect_to_sql_server_with_vba.bas
    
    
    Dim server_name As String
    Dim database_name As String
    Dim user_id As String
    Dim password As String
    
    server_name = "SCOTTWIN10-2\SQLHOME"
    database_name = "learnSQL"
    user_id = "software"
    password = "nuggets"
    
    
    Build_SQL_ConnectionString = _
        "Provider=SQLOLEDB.1;" & _
        "Data Source=" & server_name & ";" & _
        "Initial Catalog=" & database_name & ";" & _
        "User ID=" & user_id & ";" & _
        "Password=" & password & ";"
        
        
End Function



























