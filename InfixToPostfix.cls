Private theStack As StackString
Private theInput As String
Private theOutput As String

Private Sub Class_Initialize()
    
    Set theStack = New StackString
   
End Sub

Public Function Translate(target As String) As String
    
    theInput = target
    Dim inputLength As Integer
    
    inputLength = Len(theInput)
    
    For j = 1 To inputLength
        Dim ch As String
        
        ch = Mid(theInput, j, 1)
        
        Select Case ch
        
            Case "+"
                gotOper ch, 1
                
            Case "-"
                gotOper ch, 1
                
            Case "*"
                gotOper ch, 2
                
            Case "/"
                gotOper ch, 2
                
            Case "("
                theStack.push (ch)
                
            Case ")"
                gotParen ch
                
            Case Else
                theOutput = theOutput + ch
        
        End Select
        
    Next

    Do While Not theStack.isEmpty
                
        Debug.Print (theOutput)
        theOutput = theOutput + theStack.pop
                    
    Loop
    
    Translate = theOutput
    

End Function

Private Sub gotOper(opThis As String, prec1 As Integer)

    Do While Not theStack.isEmpty
        
        Dim opTop As String
        
        opTop = theStack.pop
        
        If opTop = "(" Then
            theStack.push (opTop)
            Exit Do
            
        Else
            Dim prec2 As Integer
            
            If (opTop = "+" Or opTop = "-") Then
                prec2 = 1
            
            Else
                prec2 = 2
                            
            End If
            
            If (prec2 < prec1) Then
                theStack.push (opTop)
                Exit Do
                
            Else
                theOutput = theOutput + opTop
            
            End If
            
        End If
                 
    Loop
    
    theStack.push (opThis)
    
    
End Sub

Private Sub gotParen(ch As String)

    Do While Not theStack.isEmpty
        
        Dim chx As String
        
        chx = theStack.pop
        
        If (chx = "(") Then
            Exit Do
        
        Else
            theOutput = theOutput + chx
        
        End If
    
    Loop
    
End Sub
