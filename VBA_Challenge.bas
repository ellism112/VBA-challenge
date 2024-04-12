Attribute VB_Name = "Module1"
Sub stocks()

For Each ws In Worksheets

    Dim ticker As String
    Dim volume As Double
    Dim table_row As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim max_value As Double
    Dim min_value As Double
    Dim current_cell As Double
    Dim c_cell As Double
    Dim max_vol As Double
    Dim greatest_ticker As String
    Dim smallest_ticker As String
    Dim greatest_vol_tick As String
    
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    table_row = 2
    
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    max_value = ws.Cells(2, 11).Value
    min_value = ws.Cells(2, 11).Value
    max_vol = ws.Cells(2, 12).Value
    
    
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            year_close = ws.Cells(i, 6).Value
            ticker = ws.Cells(i, 1).Value
            volume = volume + ws.Cells(i, 7).Value
            yearly_change = year_close - year_open

            ws.Range("J" & table_row).Value = yearly_change
            ws.Range("I" & table_row).Value = ticker
            ws.Range("K" & table_row).Value = (yearly_change / year_open)
            ws.Range("L" & table_row).Value = volume

            table_row = table_row + 1
            volume = 0
            
            
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            year_open = ws.Cells(i, 3).Value
            
            
        Else
            volume = volume + ws.Cells(i, 7).Value
            
        End If
        
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        ElseIf ws.Cells(i, 10).Value >= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        End If
        
        If ws.Cells(i, 11).Value < 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
            
        ElseIf ws.Cells(i, 11).Value >= 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
            
        End If
        
        ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
        
        
        current_cell = ws.Cells(i, 11).Value
        
        If current_cell > max_value Then
            max_value = current_cell
            ws.Cells(2, 16).Value = max_value
            greatest_ticker = ws.Cells(i, 9).Value
            ws.Cells(2, 15).Value = greatest_ticker
            
            
        
        ElseIf current_cell < min_value Then
            min_value = current_cell
            ws.Cells(3, 16).Value = min_value
            smallest_ticker = ws.Cells(i, 9).Value
            ws.Cells(3, 15).Value = smallest_ticker
            
        End If
            
        c_cell = ws.Cells(i, 12).Value
        
        If c_cell > max_vol Then
            max_vol = c_cell
            ws.Cells(4, 16).Value = max_vol
            greatest_vol_tick = ws.Cells(i, 9).Value
            ws.Cells(4, 15).Value = greatest_vol_tick
        
        End If
        
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
            
    Next i
    
Next ws

End Sub
