Sub Multiple_year_stock_data():

    'Loop through all sheets
    For Each ws In Worksheets
  
    'Set initial variables, total, and positions
    Dim ticker as String
    Dim yearly_change As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim percent_change As Double
    Dim total_volume As Double
        total_volume = 0
    Dim total As Integer
    Dim i As Long
    Dim row As long
        row = 2
    Dim j As Long 
        j = 2
    
    ' Create column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
       
            ' Find last row of each symbol
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        For i = 2 To LastRow
 
            ' If ticker symbol in next row is the same as the previous ticker symbol
            If ws.Range("A" & i +1).Value = ws.Range("A" & i).Value Then
                ' Add to volume in that row to the total volume
                total_volume = total_volume + ws.Range("G" & i + 1).Value
                ' print in L or "Total Stock Volume" Column
                ws.Range("L" & row).Value = total_volume 
            

            Else
                ' Print previous ticker symbol in I or "Ticker" column if 
                ticker = ws.Range("A" & i).Value
                ws.Range("I" & row).Value = ticker
    
                ' Calculate yearly change and print in J column
                open_price = ws.Range("C" & j)
                close_price = ws.Range("F" & i)
                yearly_change = (close_price - open_price)
                'print in J or "Yearly Change" column
                ws.Range("J" & row).Value = Round(yearly_change, 2)

                    'Conditional formatting for cell colors
                    ' If the value for percent change is positive, fill cells with green
                    If yearly_change > 0 Then
                        ws.Range("J" & row).Interior.ColorIndex = 4
                    ' Otherwise, fill cells with red
                    Else
                        ws.Range("J" & row).Interior.ColorIndex = 3
                    End IF
                    
                ' Calculate percent change and format for %
                If open_price <> 0 Then 
                    percent_change = yearly_change / open_price
                    ws.Range("K" & row).Value = percent_change
                    ws.Range("K" & row).NumberFormat = "0.00%"
                End If
             
                
                
                ' Reset the totals
                total_volume = 0
                row = row + 1
                j = i + 1
            
            

            
            
            End If
                
                
                
        Next i
            
    Next ws

End Sub
