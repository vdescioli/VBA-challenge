Sub stocks_all_worksheets_bonus()
    
'    Run sub on each worksheet
    For Each ws In Worksheets
    Dim ws_name As String
    ws_name = ws.Name
    
'    Variables for ticker, yearly change, %change, total,
'   summary table row and lastrow
    Dim ticker As String
    Dim yearly_change As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim table_row As Integer
    total_volume = 0
    i_i = 1
    
    table_row = 2
    
    'Column Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Loop through each row
    For i = 2 To LastRow
        
    total_volume = total_volume + Cells(i, 7).Value
        
'   Check if its the same ticker value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ticker = ws.Cells(i, 1).Value
        
        ws.Range("I" & table_row).Value = ticker
    
'    To determine yearly change value
               
        i_i = i_i + 1
        
        open_price = ws.Cells(i_i, 3).Value
                
        close_price = ws.Cells(i, 6).Value
        
        yearly_change = close_price - open_price
        ws.Range("J" & table_row).Value = yearly_change
        ws.Range("J" & table_row).NumberFormat = "$0.00"
        
'    Color Formatting
            If ws.Range("J" & table_row).Value > 0 Then
            ws.Range("J" & table_row).Interior.ColorIndex = 4
            Else
            ws.Range("J" & table_row).Interior.ColorIndex = 3
            End If
            
        yearly_change = 0
        
'    Output total value and then reset
                                   
        ws.Range("L" & table_row).Value = total_volume
        
        total_volume = 0
        
'    Find percent change, output value, and reset
    
            If open_price = 0 Then
            percent_change = close_price
            
            Else
            yearly_change = close_price - open_price
            percent_change = yearly_change / open_price
            
            End If
            
        ws.Range("K" & table_row).Value = percent_change
        ws.Range("K" & table_row).NumberFormat = "0.00%"
        percent_chane = 0
        
        table_row = table_row + 1
        i_i = i
        
        End If
    
    Next i
    
'    Bonus !!
'   Variables
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_vol As Double
    
    'Rows & Headers
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
'    Find last row in summary table
    LastRow_2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    max_increase = 0
    max_decrease = 0
    max_vol = 0
    
'    Run Loop of summary table

    For j = 2 To LastRow_2
    
'        Greatest Increase %
        If ws.Cells(j, 11).Value > max_increase Then
        max_increase = ws.Cells(j, 11).Value
        ws.Range("O2").Value = ws.Range("I" & j).Value
        ws.Range("P2").Value = ws.Range("K" & j).Value
        ws.Range("P2").NumberFormat = "0.00%"
        
        
        Else
        max_increase = max_increase
        
        End If
        
'        Greatest Decrease %
        If ws.Cells(j, 11).Value < max_decrease Then
        max_decrease = ws.Cells(j, 11).Value
        ws.Range("O3").Value = ws.Range("I" & j).Value
        ws.Range("P3").Value = ws.Range("K" & j).Value
        ws.Range("P3").NumberFormat = "0.00%"
        
        Else
        max_decrease = max_decrease
        
        End If
        
'        Greatest stock volume
        If ws.Cells(j, 12).Value > max_vol Then
        max_vol = ws.Cells(j, 12).Value
        ws.Range("O4").Value = ws.Range("I" & j).Value
        ws.Range("P4").Value = ws.Range("L" & j).Value
        
        Else
        max_vol = max_vol
        
        End If
   
   
   Next j

    
   Next ws
End Sub


