Attribute VB_Name = "String"
Option Compare Database

Public Function RemoveSpaces(text As String) As String
    
    RemoveSpaces = Replace(text, " ", "")
    
End Function
