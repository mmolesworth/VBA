Public Sub FormatWorksheet()

    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    
    Dim wsCount As Integer
    Dim wsCounter As Integer
    
    Set wb = ActiveWorkbook
    wsCount = wb.Worksheets.Count
    wsCounter = 1
    
    Do While wsCounter <= wsCount
    
        Set ws = wb.Worksheets(wsCounter)
        ws.Activate
        
        Cells.Select
        
        'add filter
        If ws.AutoFilterMode = False Then
            Selection.AutoFilter
        
        End If
        
        'freeze top row
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        
        ActiveWindow.FreezePanes = True
        
        'set column widths
        ws.Range("A:A").EntireColumn.ColumnWidth = 13
        ws.Range("B:B").EntireColumn.ColumnWidth = 13
        ws.Range("C:C").EntireColumn.ColumnWidth = 13
        ws.Range("D:D").EntireColumn.ColumnWidth = 13
        ws.Range("E:E").EntireColumn.ColumnWidth = 30
        ws.Range("F:F").EntireColumn.ColumnWidth = 55
        ws.Range("G:G").EntireColumn.ColumnWidth = 55
        ws.Range("H:H").EntireColumn.ColumnWidth = 30
        ws.Range("I:I").EntireColumn.ColumnWidth = 30
        ws.Range("J:J").EntireColumn.ColumnWidth = 30
        
        
        'set row heights
        Cells.Select
        Cells.EntireRow.AutoFit
    
    
        'move counter forward
        wsCounter = wsCounter + 1
        
    Loop
    
End Sub
