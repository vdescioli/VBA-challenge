Sub stocks_single_worksheet()

'Column Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
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
    LastRow = Cells(Rows.Count, 1).End(xlUp).row
    table_row = 2
    
'Loop through each row
    For i = 2 To LastRow
        
    total_volume = total_volume + Cells(i, 7).Value
        
'   Check if its the same ticker value
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ticker = Cells(i, 1).Value
        
        Range("I" & table_row).Value = ticker
    
'    To determine yearly change value
               
        i_i = i_i + 1
        
        open_price = Cells(i_i, 3).Value
                
        close_price = Cells(i, 6).Value
        
        yearly_change = close_price - open_price
        Range("J" & table_row).Value = yearly_change
        Range("J" & table_row).NumberFormat = "$0.00"
        
'    Color Formatting
            If Range("J" & table_row).Value > 0 Then
            Range("J" & table_row).Interior.ColorIndex = 4
            Else
            Range("J" & table_row).Interior.ColorIndex = 3
            End If
            
        yearly_change = 0
        
'    Output total value and then reset
                                   
        Range("L" & table_row).Value = total_volume
        
        total_volume = 0
        
'    Find percent change, output value, and reset
    
            If open_price = 0 Then
            percent_change = close_price
            
            Else
            yearly_change = close_price - open_price
            percent_change = yearly_change / open_price
            
            End If
            
        Range("K" & table_row).Value = percent_change
        Range("K" & table_row).NumberFormat = "0.00%"
        percent_chane = 0
        
        table_row = table_row + 1
        i_i = i
        
        End If
    
    Next i
    

End Sub
