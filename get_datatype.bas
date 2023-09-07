
Function GetDataType(var As Variant) As String

    Select Case typeName(var)
    Case "Variant"
        GetDataType = "Variant"
    Case "Integer"
        GetDataType = "Integer"
    Case "Long"
        GetDataType = "Long"
    Case "Single"
        GetDataType = "Single"
    Case "Double"
        GetDataType = "Double"
    Case "Currency"
        GetDataType = "Currency"
    Case "String"
        GetDataType = "String"
    Case "Object"
        GetDataType = "Object"
    Case "Boolean"
        GetDataType = "Boolean"
    Case "Date"
        GetDataType = "Date"
    Case "Error"
        GetDataType = "Error"
    Case "Empty"
        GetDataType = "Empty"
    Case "Null"
        GetDataType = "Null"
    Case "Byte"
        GetDataType = "Byte"
    Case "User-Defined Type"
        GetDataType = "User-Defined Type"
    Case "Array"
        GetDataType = "Array"
    Case Else
        GetDataType = "Unknown"
    End Select
    
End Function

Sub main()
    Dim a As Date, b As Currency, c As String
    Debug.Print GetDataType(a)
    Debug.Print GetDataType(b)
    Debug.Print GetDataType(c)
End Sub
