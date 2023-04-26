Sub main()

    '->Tools->References->
    'Microsoft ActiveX Data Object  2.8 Library
    
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
   
    
    'Dim conn As New ADODB.Connection
    Dim server_name As String
    Dim database_name As String
    Dim user_id As String
    Dim password As String
    
    
    server_name = "SCOTTWIN10-2\SQLHOME"
    database_name = "learnSQL"
    user_id = "software"
    password = "nuggets"
    
    
    conn.ConnectionString = _
        "Provider=SQLOLEDB.1;" & _
        "Data Source=" & server_name & ";" & _
        "Initial Catalog=" & database_name & ";" & _
        "User ID=" & user_id & ";" & _
        "Password=" & password & ";"
        
    conn.Open
    
    On Error GoTo CloseConnection


    If conn.State <> adStateOpen Then
        Debug.Print "invalid connection string"
        Debug.Print conn.ConnectionString
        End
    End If
    
CloseConnection:
    conn.Close
    Set conn = Nothing
End Sub
