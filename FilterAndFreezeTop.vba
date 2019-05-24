Public Sub FilterAndFreezeTop(ws As Worksheet)
'Format worksheet (auto-filter & freeze top row)

    ws.UsedRange.AutoFilter
    ws.Activate

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    
    ws.Rows("1:1").Select
    ActiveWindow.FreezePanes = True
    
    
End Sub
