Option Compare Database

Private variables() As Variant
Private isLoaded As Boolean


Public Function GetCurrentUsername() As String
    
    GetCurrentUsername = (Environ$("Username"))
    
    
End Function

Public Sub LoadEnvVariables()
    
    Erase variables()
    ReDim variables(1000)
    
    Dim i As Integer
    
    i = 1
    
    Do
        variables(i - 1) = Split(Environ(i), "=")
        i = i + 1
        
    Loop Until Environ(i) = ""
    
    isLoaded = True
    
    
End Sub

Public Function GetVariableNames() As String()
    
    If isLoaded = False Then
        LoadEnvVariables
    
    End If
    
    Dim names(1000) As String
    Dim i As Integer
    
    i = 0
    
    Do Until IsEmpty(variables(i)) = True
        names(i) = variables(i)(0)
        i = i + 1
        
    Loop
    
    GetVariableNames = names
    

End Function

Public Function GetVariableValue(name As String) As String
    
    GetVariableValue = (Environ$(name))


End Function
Public Sub PrintVariablesToImmediate()
    
    Dim i As Integer
    
    i = 1
    
    Do
        Debug.Print (Environ(i))
        i = i + 1
        
    Loop Until Environ(i) = ""
    
    
End Sub
