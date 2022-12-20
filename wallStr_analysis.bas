Attribute VB_Name = "Module1"
Sub wallst()
    Dim i, j As Double
    Dim k, n As Double
    Dim cur_total As Double
    'change
    'open at earliest/lowest date
    Dim ld_open, ld As Single
    'close at latest/largest date
    Dim hd_close, hd As Single
    'change
    Dim y_change As Single
    
    For Each ws In Worksheets
    
        'clear entries to avoid errors
        ws.Range("I:Z").ClearContents
    
        
        cur_total = 0
        'count how many rows, assign it to var so it can be used in for loop (this allows list size to be variable)
        k = ws.Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count
        
        'populate nonduplicate list
        ws.Range(ws.Cells(1, 1), ws.Cells(k, 1)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("i1"), Unique:=True
        
        'count the number of cells in new table for for loop
        n = ws.Range("i:i").Cells.SpecialCells(xlCellTypeConstants).Count
        
       ' MsgBox Range("A1").End(xlDown).Row
        
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Range("I1:L1").Columns.AutoFit
        
        For i = 2 To n
            cur_total = 0
            ld = 0
            hd = 0
            ld_open = 0
            hd_close = 0
            'MsgBox cur_total
            For j = 2 To k
                If ws.Cells(i, 9).Value = ws.Cells(j, 1).Value Then
            cur_total = cur_total + ws.Cells(j, 7).Value
                    
                    If ld = 0 Or hd = 0 Then
                        ld = ws.Cells(j, 2).Value
                        hd = ws.Cells(j, 2).Value
                        ld_open = ws.Cells(j, 3).Value
                        hd_close = ws.Cells(j, 6).Value
                        
                        'MsgBox (Cells(j, 2).Value)
                    End If
                    
                    If ld > ws.Cells(j, 2).Value Then
                        ld = ws.Cells(j, 2).Value
                        ld_open = ws.Cells(j, 3).Value
                    End If
                    
                    If hd < ws.Cells(j, 2).Value Then
                        hd = ws.Cells(j, 2).Value
                        hd_close = ws.Cells(j, 6).Value
                    End If
                End If
                'MsgBox (Cells(j, 1).Value)
                'MsgBox (Cells(i, 9).Value)
            Next j
            
            'calculate yearly change
            y_change = hd_close - ld_open
            
            'Cells(i, 10).Value = y_change
            ws.Cells(i, 10).Value = Format(y_change, "#.00")
            If y_change <= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
            
            ws.Cells(i, 11).Value = FormatPercent((y_change / ld_open))
            ws.Cells(i, 12).Value = cur_total
        Next i
        
        
        'extras
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(2, 17).Value = FormatPercent(WorksheetFunction.Max(ws.Range("k:k")))
        'Cells(2, 16).Value = WorksheetFunction.VLookup(Cells(2, 17).Value, Sheet1.Range("k:k"), -2, False)
        
        
        ws.Cells(3, 15).Value = "Greatest % decreaase"
        ws.Cells(3, 17).Value = FormatPercent(WorksheetFunction.Min(ws.Range("k:k")))
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("l:l"))
        
        'find matching ticker
        
        For i = 2 To n
            If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i
        
        For i = 2 To n
            If ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i
        
        ws.Range("o:q").Columns.AutoFit
         
        'MsgBox (k)
        'MsgBox (n)
    Next ws

End Sub


