Public Sub ParseRule()
    
    Dim source As Worksheet
    Dim analysis As Worksheet
    Dim sourceCell As Range
    Dim destinationCell As Range
    Dim tokens() As String
    Dim stack As StackString
    Dim conclusion As String
    Dim condition As String
    Dim stringBuilder As New stringBuilder
    Dim temp As String
    Dim newStack As New StackString
    Dim tokens2() As String
    Dim stack2 As New StackString
    Dim newStack2 As New StackString
    
    Dim count As Integer
        
    count = 1
        
    Set source = ActiveWorkbook.Worksheets("source")
    Set destination = ActiveWorkbook.Worksheets("analysis")
    
    Set sourceCell = source.Range("A2")
    Set destinationCell = destination.Range("A2")
    
    'loop through rules
    Do Until sourceCell.value = ""
        
        count = count + 1
        
        'Copy rule meta-data
        destinationCell.value = sourceCell.Offset(0, 6).value                'TBDID
        destinationCell.Offset(0, 1).value = sourceCell.Offset(0, 7).value   'RULEID
        destinationCell.Offset(0, 2).value = sourceCell.Offset(0, 8).value   'Interface
        destinationCell.Offset(0, 3).value = Trim(ReplaceCommonPhrases(sourceCell.Offset(0, 13).value)) 'Simplified Rule Text
        
        'Tokenize rule text string
        tokens = Split(destinationCell.Offset(0, 3).value, " ")
        
        stringBuilder.Clear
        Set stack = New StackString
        Set newStack = New StackString
        Set stack2 = New StackString
        Set newStack2 = New StackString
        
        Dim sb As New stringBuilder

        Debug.Print ("Row: " & count)
        
        For Each token In tokens
            
            
            
            Select Case Trim(token)
                
                Case "NOT-NULL"
                    temp = stack.pop
                    stack.push (token)
                    stack.push (temp)

                Case Else
                    stack.push (token)
                    
            End Select
            
        Next

        'Reverse stack

        
        Do Until stack.isEmpty
            newStack.push (stack.pop)
            
        Loop

        Do Until newStack.isEmpty
            sb.Append (newStack.pop)
            sb.Append (" ")
        
        Loop
        
        'Debug.Print (sb.Text)
        
        
        
        'PART II
        
        
        
        tokens2 = Split(sb.Text, " ")
        
        For Each token In tokens2
            
            Select Case Trim(token)
                
                Case " "
                    'ignore
                    
                Case "NOT-NULL"
                    If Not stack2.isEmpty Then
                        stack2.push (Chr(41))
                        stack2.push (Chr(41))
                        stack2.push (" ")
                    
                    End If
                    
                    stack2.push ("(CONCLUSION (")
                    stack2.push ("NOT-NULL ")
                    
                Case "IF"
                    If Not stack2.isEmpty Then
                        stack2.push (Chr(41))
                        stack2.push (Chr(41))
                        stack2.push (" ")
                        
                    End If
                    
                    stack2.push ("(CONDITIONS ")
                    stack2.push ("(")
                    
                Case "-"
                    
                    
                Case "."
                    stack2.push (Chr(41))
                    
                Case Else
                    stack2.push (Trim(token))
                    stack2.push (" ")
            
            End Select
            
        Next
        
        sb.Clear
        
        Do Until stack2.isEmpty
            newStack2.push (stack2.pop)
        
        Loop
        
        Do Until newStack2.isEmpty
            sb.Append (newStack2.pop)
            
        Loop
        
        destinationCell.Offset(0, 4).value = "(" & sb.Text & Chr(41)
        
        sb.Clear
        
        
        
        
        
        
        
        
        'Write conclusion/condition to Sheet2
        
        Set destinationCell = destinationCell.Offset(1, 0)
        Set sourceCell = sourceCell.Offset(1, 0)

    Loop
    
    

End Sub

Private Function ReplaceCommonPhrases(value As String) As String
    
    Dim newValue As String
    
    newValue = value
    
    newValue = Replace(newValue, Chr(10), "")
    newValue = Replace(newValue, Chr(13), "")
    newValue = Replace(newValue, Chr(34), "")
    newValue = Replace(newValue, ",", "")
    newValue = Replace(newValue, ".", " . ")
    newValue = Replace(newValue, "a submitted", "")
    newValue = Replace(newValue, "is the issuer", "")
    newValue = Replace(newValue, "must be populated", "NOT-NULL")
    newValue = Replace(newValue, "if all of the following is true:", "IF")
    newValue = Replace(newValue, "when", "IF")
    newValue = Replace(newValue, "When", "IF")
    newValue = Replace(newValue, " if ", " IF")
    newValue = Replace(newValue, "is equal to", "=")
    newValue = Replace(newValue, " is ", " = ", , , vbTextCompare)
    newValue = Replace(newValue, "indicates", "=")
    newValue = Replace(newValue, "the", "")
    newValue = Replace(newValue, "The", "")
    
    ReplaceCommonPhrases = newValue
    

End Function

Function Reverse(Text As String) As String
    Dim i As Integer
    Dim StrNew As String
    Dim strOld As String
    strOld = Trim(Text)
    For i = 1 To Len(strOld)
      StrNew = Mid(strOld, i, 1) & StrNew
    Next i
    Reverse = StrNew
End Function

Function ReverseStringArray(arr() As String) As String()

    Dim newArray(100) As String
    Dim newIndex As Integer
    
    newIndex = 0
    
    For i = UBound(arr) To 0 Step -1
        
        newArray(newIndex) = arr(i)
        
        newIndex = newIndex + 1
        
    Next
    
    
    ReverseStringArray = newArray
    
End Function


