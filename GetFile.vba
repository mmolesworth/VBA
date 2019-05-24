Private Function GetFile(directory As String, pattern As String) As file
    
'External libraries'
    'FileSystem objects are found in the Microsoft Scripting Runtime
    'Tools > References | Ensure 'Microsoft Scripting Runtime' is checked
    
    'Regular Expression objects are found in the Microsoft VBScipt Regular Expression 5.5 library
    'Tools > References | Ensure 'Microsoft VBScipt Regular Expression 5.5' is checked

'Function Objective'
    'Search through the current file directory to see if a file exists that matches the Decion output format
    'If no file exists, prompt user to add one
    
    Dim fso As New FileSystemObject
    Dim myFolder As Folder
    Dim aFile As file
    Dim regex As New RegExp
    Dim fileNamePattern As String: fileNamePattern = pattern
    Dim targetFile As file
    
    Set myFolder = fso.GetFolder(directory)
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = fileNamePattern
        
    End With
    
    'Loop through the files in the current directory to see if the Decision file already exists
    For Each aFile In myFolder.Files
        
        If regex.Test(aFile.Name) Then
            Set targetFile = aFile
'#TODO stop loop once it's found

        End If
        
    Next
    
    'If decisionFile Is Null Then
        '#TODO prompt user to input filename
        'Debug.Print ("Decision file not found.")
        
    'End If
    
    Set GetFile = targetFile
    
    
End Function
