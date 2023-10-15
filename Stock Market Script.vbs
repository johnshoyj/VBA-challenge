Sub stock_market()

Dim SummaryTableRow As Long
Dim ticker As String
Dim stock_open As Double
Dim stock_close As Double
Dim volume As Double
Dim yearly_chg As Double
Dim pct_chg As Double
Dim lg_increase As Double
Dim lg_decrease As Double
Dim lg_volume As Double
Dim lg_increase_ticker As String
Dim lg_decrease_ticker As String
Dim lg_volume_ticker As String

    For Each ws In Worksheets
        volume = 0
        stock_open = 0
        stock_close = 0
        yearly_chg = 0
        pct_chg = 0
        lg_increase = 0
        lg_decrease = 0
        lg_volume = 0
        SummaryTableRow = 2
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
        
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
                If i > 2 Then
                    
                    stock_close = ws.Cells(i - 1, 6).Value
                    yearly_chg = stock_close - stock_open
                    
                    If yearly_chg < 0 Then
                        ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                    End If
                    
                    If stock_open <> 0 Then
                        pct_chg = (stock_close - stock_open) / stock_open
                    Else
                        pct_chg = 0
                    End If
                    
                    If pct_chg > lg_increase Then
                        lg_increase = pct_chg
                        lg_increase_ticker = ticker
                    End If
                    
                    If pct_chg < lg_decrease Then
                        lg_decrease = pct_chg
                        lg_decrease_ticker = ticker
                    End If
                    
                    If volume > lg_volume Then
                        lg_volume = volume
                        lg_volume_ticker = ticker
                    End If
                    
                    ws.Cells(SummaryTableRow, 9).Value = ticker
                    ws.Cells(SummaryTableRow, 10).Value = yearly_chg
                    ws.Cells(SummaryTableRow, 11).Value = pct_chg
                    ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
                    ws.Cells(SummaryTableRow, 12).Value = volume
        
                    SummaryTableRow = SummaryTableRow + 1
                
                End If
                
                ticker = ws.Cells(i, 1).Value
                volume = 0
                stock_open = ws.Cells(i, 3).Value
            
            End If
                
            volume = volume + ws.Cells(i, 7).Value
        
        Next i
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Columns("N:P").AutoFit
        
        ws.Cells(2, 15).Value = lg_increase_ticker
        ws.Cells(3, 15).Value = lg_decrease_ticker
        ws.Cells(4, 15).Value = lg_volume_ticker
        
        ws.Cells(2, 16).Value = lg_increase
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = lg_decrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = lg_volume
        ws.Cells(4, 16).NumberFormat = "0000000000000"
        
   Next ws

End Sub
