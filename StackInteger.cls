Option Explicit
Private maxSize As Integer
Private stackArray(1 To 100) As Integer
Private top As Integer

Private Sub Class_Initialize()

    maxSize = 100
    top = 1
       
   
End Sub

Private Sub Class_Terminate()
    
    maxSize = 0
    top = 0
    Erase stackArray
    
    
End Sub

Public Sub push(j As Integer)
    
    stackArray(top) = j
    
    top = top + 1
    

End Sub

Public Function pop() As Integer
    
    Dim x As Integer
    
    x = stackArray(top - 1)
    
    top = top - 1
    
    pop = x
    
End Function

Public Function peek() As Integer

    peek = stackArray(top - 1)
    
End Function

Public Property Get isEmpty() As Boolean
    
    If top = 1 Then
        isEmpty = True
        
    Else
        isEmpty = False
        
    End If
    
End Property

Public Property Get isFull() As Boolean
    
    If top = maxSize Then
        isFull = True
        
    Else
        isFull = False
        
    End If
    
End Property




