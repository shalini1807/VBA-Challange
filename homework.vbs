Sub Year_stock()
    Dim ticker As String
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim Lastrow As Long
    Dim year_open As Double, year_close As Double
    Dim Total_Stock_Volume As Double
    Dim previous_i As Long
    Dim output_row As Long

    'Loop through each All sheets
    For Each ws In Worksheets
        'Assign a column header for every task we are going perform
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent change"
        ws.Range("L1").Value = "Total stock Volume"
     
        'Assign integer for the loop to start
        output_row = 2
        Total_Stock_Volume = 0
        previous_i = 2
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
     
        'for each ticker summarize through loop quarterly change,percent change and total stock volume
        For i = 2 To Lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
        
                'Get the value first day open from the column 3 or "C" and last day close of the year on column 6 or "F"
                year_open = ws.Cells(previous_i, 3).Value
                year_close = ws.Cells(i, 6).Value
        
                'Calculate quarterly change and percent change
                quarterly_change = year_close - year_open
                percent_change = If(year_open <> 0, quarterly_change / year_open, 0)
        
                'Sum the total stock volume
                For j = previous_i To i
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
                Next j
        
                'Output results
                ws.Cells(output_row, 9).Value = ticker
                ws.Cells(output_row, 10).Value = quarterly_change
                ws.Cells(output_row, 11).Value = percent_change
                ws.Cells(output_row, 12).Value = Total_Stock_Volume
        
                'Check for greatest values
                
                If percent_change > max_increase Then
                    max_increase = percent_change
                    max_increase_ticker = ticker
                ElseIf percent_change < max_decrease Then
                    max_decrease = percent_change
                    max_decrease_ticker = ticker
                End If
                
                If Total_Stock_Volume > max_volume Then
                    max_volume = Total_Stock_Volume
                    max_volume_ticker = ticker
                End If
        
                'Reset for next ticker
                output_row = output_row + 1
                Total_Stock_Volume = 0
                previous_i = i + 1
            End If
        Next i
        
        ' Output greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 16).Value = max_increase_ticker
        ws.Cells(2, 17).Value = max_increase
        ws.Cells(3, 16).Value = max_decrease_ticker
        ws.Cells(3, 17).Value = max_decrease
        ws.Cells(4, 16).Value = max_volume_ticker
        ws.Cells(4, 17).Value = max_volume
        
        ' Format percent changes
        ws.Range("K:K,Q2:Q3").NumberFormat = "0.00%"
        ws.Range("L:L,Q4").NumberFormat = "#,##0"
    Next ws
    MsgBox ("Process complete")
End Sub

