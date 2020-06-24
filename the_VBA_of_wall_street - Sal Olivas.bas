Attribute VB_Name = "Module11"
Sub the_vba_of_wall_stree()

For Each ws In Worksheets

    Dim WorksheetName As String
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    WorksheetName = ws.Name
    
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim stock_volume As Double
    Dim start_open As Long
    
    'setting variables
    ws.Range("I1").Value = "ticker"
    ws.Range("J1").Value = "yearly_change"
    ws.Range("K1").Value = "percent_change"
    ws.Range("L1").Value = "stock_volume"

    stock_volume = 0
    
    ' utilizing last row func.
    Dim summary_row As Long
    summary_row = 2
    
    
    '**firstday of month open
    start_open = 2
    
    'begin for loop
    For i = 2 To LastRow
        'checking that next cell is different than current cell
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'establishing ticker
            ticker = Cells(i, 1).Value
            
            'adding to the stock volume
            stock_volume = stock_volume + Cells(i, 7).Value
            
            'adding into summary table
            ws.Range("I" & summary_row).Value = ticker
            
            ws.Range("L" & summary_row).Value = stock_volume
            
            'establishing yearly_open & yearly_close
            yearly_open = ws.Cells(start_open, 3)
            
            yearly_close = ws.Cells(i, 6)
            
            yearly_change = yearly_close - yearly_open
                
                'accounting for 0 on yearly_open
                On Error Resume Next
                
                If percent_change = 0 Then
                    Resume Next
                    
                End If
            
            'establishing percent_change
            percent_change = Round(((yearly_close - yearly_open) / yearly_open) * 100, 2)
                
                'NEED TO ADD IF STATEMENT HERE FOR DIVIDE BY 0 ISSUE
               
            'adding into summary table
            ws.Range("J" & summary_row).Value = yearly_change
            
            ws.Range("K" & summary_row).Value = "%" & percent_change
            
            'assigning color inex
                If ws.Range("J" & summary_row).Value >= 0 Then
            
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4
                
                Else
                
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3
                    
                End If
            
            'adding 1 to the summary table row
            summary_row = summary_row + 1
            
            start_open = i + 1
            
            'resetting the total volume
            stock_volume = 0
            
            Else
                ' Add to the Brand Total
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
            
            End If
            
        Next i
    
Next ws
    
End Sub

