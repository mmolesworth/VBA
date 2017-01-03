Public Function CountInterfaces(text As String) As Integer

    Dim textArray() As String
    
    textArray = Split(text, ";")
    
    CountInterfaces = UBound(textArray) + 1
    
    
End Function
