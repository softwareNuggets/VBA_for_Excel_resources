Attribute VB_Name = "Module2"
Sub LoadFoldersIntoSheet()

    Dim ws As Worksheet
    Dim startingPath As String
    Dim folderList As Variant
    Dim i As Integer
    Dim FolderSize As Long
    
    Set ws = Worksheets("Sheet1")
    
    ' Specify the root folder path
    startingPath = "D:\YouTube"
    
    ' Get list of folders
    folderList = GetRootFolders(startingPath)
    
    If IsArray(folderList) Then
    
        For i = LBound(folderList) To UBound(folderList)
        
            ' Writing folder names to column A
            ws.Cells(i + 1, 1).Value = folderList(i)
            
            ' Build full path
            Dim folderName As String
            folderName = startingPath + "\" + folderList(i)
            
            ' Writing folder filesize to column B
            ws.Cells(i + 1, 2).Value = GetFolderSize(folderName)
            
        Next i
    Else
        MsgBox "No subfolders found or invalid folder path.", vbExclamation
    End If
    
    Set ws = Nothing ' Release the reference to the worksheet
    folderList = Nothing ' Clear the array
    
End Sub

Function GetRootFolders(ByVal rootPath As String) As Variant
    'return an array of folder names
    
    Dim fso As Object
    Dim folder As Object
    Dim folderList() As String
    Dim count As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(rootPath) Then
        GetRootFolders = Array()
        Exit Function
    End If
    
    Set folder = fso.GetFolder(rootPath)
    
    ReDim folderList(folder.SubFolders.count - 1)
    count = 0

    For Each subFolder In folder.SubFolders
        folderList(count) = subFolder.Name
        count = count + 1
    Next subFolder
    
    GetRootFolders = folderList
    
    Set folder = Nothing
    Set fso = Nothing
    
End Function

Function GetFolderSize(ByVal sPath As String) As Double
    Dim fso As Object
    Dim folder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(sPath) Then
        Set folder = fso.GetFolder(sPath)
        GetFolderSize = folder.Size
    Else
        GetFolderSize = 0
    End If
    
    Set folder = Nothing
    Set fso = Nothing
    
End Function
