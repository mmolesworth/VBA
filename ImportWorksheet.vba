Public Function ImportWorksheet(destination As Workbook, importFile As file, _
                                            destinationWorksheetLabel As String, _
                                            Optional worksheetIndex As Integer = 1) As Worksheet

'External libraries'
    'Regular Expression objects are found in the Microsoft VBScipt Regular Expression 5.5 library
    'Tools > References | Ensure 'Microsoft VBScipt Regular Expression 5.5' is checked
    
'Function Objective:
    'Copy the data contained in the import file into the specifie workbook as a new worksheet.
    'Provide an auto-filter and freeze the top row of the worksheet.

    Dim sourceWorkbook As Workbook
    Dim newWorksheet As Worksheet
    
    
    
'Open original in the background
    If importFile.Name <> "" Then
        Set sourceWorkbook = Application.Workbooks.Open(importFile.Path)
        sourceWorkbook.Application.Visible = True
        
    Else
    '#TODO Handle case when file is not found and end process
            
    End If

'Remove autofilter, if present
sourceWorkbook.Worksheets(worksheetIndex).AutoFilterMode = False


'Copy contents into new worksheet
    Set newWorksheet = destination.Worksheets.Add()
    newWorksheet.Name = destinationWorksheetLabel
    
    sourceWorkbook.Worksheets(worksheetIndex).UsedRange.copy destination:=newWorksheet.Range("A1")

'Close original worksheet
    sourceWorkbook.Close

    Set ImportWorksheet = newWorksheet
    
End Function
