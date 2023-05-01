Sub treasury_gov()

   ' add reference Microsoft WinHTTP Services
    Dim ws As Worksheet
    Set ws = Worksheets("sheet1")
    
    Dim stocks As String
    Dim flags As String
    Dim Url As String
    Dim results As String
    Dim baseURL As String
    Dim service As String
    Dim params As String
    
    Dim fields As String
    Dim filter As String
    Dim format As String
    
    fields = "record_date,tot_pub_debt_out_amt"
    filter = "record_date:gte:2023-04-01"
    format = "csv"
    
    'url
    'https://fiscaldata.treasury.gov/datasets/debt-to-the-penny/debt-to-the-penny
    
    baseURL = "https://api.fiscaldata.treasury.gov/services/api/fiscal_service"
    service = "/v2/accounting/od/debt_to_penny"
    params = "?fields=" & fields & "&filter=" & filter & "&format=" & format
    
    
    Url = baseURL & service & params
    Debug.Print Url
    
    Dim Http As New WinHttpRequest
    Http.Open "GET", Url, False
    Http.Send
    
    'Debug.Print Http.ResponseText
    results = Http.ResponseText
    
    
    Call ReadLinebyLine(results)
    
    MsgBox ("All done")
    
    
End Sub


Sub ReadLinebyLine(inputData As String)
    ' Split the input data into an array of lines
   
   Dim ws As Worksheet
   Set ws = Worksheets("Sheet1")
   
    ws.Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy"
    
    ws.Columns("B:B").Select
    Selection.NumberFormat = "$#,##0.00"
    
    
    Dim lines() As String
    lines = Split(inputData, Chr$(10))
   
    Dim row As Integer
    row = 1
    Dim j As Integer
    For j = 1 To UBound(lines) - 1
        Dim values() As String
        values = Split(lines(j), ",")
        
        d = Replace(values(0), Chr(34), "")
        n = Replace(values(1), Chr(34), "")
        ws.Cells(row, 1) = d
        
        ws.Cells(row, 2).Value = n
        row = row + 1
    Next j
    
End Sub