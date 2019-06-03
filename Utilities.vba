Public Sub DebugPaste(sep As String, ParamArray values() As Variant)

    Dim value As Variant
    Dim outputString As String
    Dim index As Integer
    Dim upperBound As Integer
    
    outputString = ""
    upperBound = UBound(values())
    
    For index = 0 To upperBound
        
        If index < upperBound Then
            outputString = outputString & CStr(values(index)) & sep
            
        Else
            outputString = outputString & CStr(values(index))
        
        End If
        
    Next
    
    Debug.Print (outputString)
    
    
End Sub

Public Function Paste(sep As String, ParamArray values() As Variant) As String

    Dim value As Variant
    Dim outputString As String
    Dim index As Integer
    Dim upperBound As Integer
    
    outputString = ""
    upperBound = UBound(values())
    
    For index = 0 To upperBound
        
        If index < upperBound Then
            outputString = outputString & CStr(values(index)) & sep
            
        Else
            outputString = outputString & CStr(values(index))
        
        End If
        
    Next
    
    Paste = outputString
    
End Function

Public Function PasteCollapseRange(r As Range, sep As String) As String
    
    Dim cell As Range
    Dim outputString As String
    Dim upperBound As Integer
    Dim index As Integer
    
    outputString = ""
    upperBound = r.Rows.Count
    
    For index = 1 To upperBound
        
        If index < upperBound Then
            outputString = outputString & CStr(r.Cells(index, 1).Text) & sep
            
        Else
            outputString = outputString & CStr(r.Cells(index, 1).Text)
        
        End If
        
    Next
    
    PasteCollapse = outputString
    
End Function

Public Function PasteCollapseArray(arr() As String, sep As String) As String
    
    Dim outputString As String
    Dim upperBound As Integer
    Dim index As Integer
    
    outputString = ""
    upperBound = UBound(arr)
    
    For index = 0 To upperBound
        
        If index < upperBound Then
            outputString = outputString & CStr(arr(index)) & sep
            
        Else
            outputString = outputString & CStr(arr(index))
        
        End If
        
    Next
    
    PasteCollapseArray = outputString
    
End Function

Public Sub ListWorksheets(wb As Workbook)

    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        DebugPaste "-", ws.index, ws.name
        
    Next
    
End Sub

Public Function GetWorksheetByName(wb As Workbook, name As String) As Worksheet
    
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        
        If ws.name = name Then
            Set GetWorksheetByName = ws
            Exit Function
            
        End If
        
    Next
        
    Set GetWorksheetByName = Null
    
End Function

Public Sub test()
    ListWorksheets ActiveWorkbook
    
    Debug.Print (GetWorksheetByName(ActiveWorkbook, "CSS-Holidays").index)
    
    

End Sub
