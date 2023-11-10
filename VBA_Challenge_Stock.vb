Sub Multiple_year_stock_year():


    
    Dim i, j As Integer
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    'To create Ticker,Yearly Change,Percent Change & Total Stock Volume (Columns)
    
        volume_counter = 0
        Ticker = 2
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Set myrange = ws.Range("A:C")
        
     'To assign the Ticker, Value (columns) & Greatest % increase, Greatest % decrease, Greatest Total volume (values in the columns), respectively!
     
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest Total volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
    'Looping function & Vlookup function performed to populate the Ticker, Yearly chnage, Percent Change & Total Stock Volume respectively!
    
        For i = 2 To ws.Range("A1").End(xlDown).Row
        
        
            If (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
                
                volume_counter = volume_counter + ws.Cells(i, 7).Value
                
            Else
                ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - Application.WorksheetFunction.VLookup(ws.Cells(Ticker, 9).Value, myrange, 3, False)
                ws.Cells(Ticker, 11).Value = (ws.Cells(Ticker, 10).Value / Application.WorksheetFunction.VLookup(ws.Cells(Ticker, 9).Value, myrange, 3, False)) * 100
                
                ws.Cells(Ticker, 12).Value = volume_counter
                volume_counter = volume_counter + ws.Cells(i, 7).Value
                
                
                volume_counter = 0
                Ticker = Ticker + 1
        
            End If
            
        Next i
    
    'Color Index in the respective columns!
    
        For j = 2 To ws.Range("J1").End(xlDown).Row
            If ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
                
            Else:
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            
            End If
        Next j
    
    'To populate and match the Max, Min and Total volume (values) in the respective columns!
    
              ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("k:k"))
              ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("k:k"))
              ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
              
              ws.Range("P2").Value = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K:K"), 0))
              ws.Range("P3").Value = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K"), 0))
              ws.Range("P4").Value = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L:L"), 0))
              
              
    'Made appropriate adjustments to VBA script to enable it to run on every worksheet (that is, every year) at once through "For Each ws In Worksheets"
    
    Next ws
    
               
End Sub
