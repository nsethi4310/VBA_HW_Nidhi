Sub tickser_summary()

  ' Setting  initial variables for holding the ticker name, worksheet, vol_total, summary_table, yearly_chnage, yrl_chg_pct, open_price, close_price and last row
    Dim ticker_Name As String
    Dim ws As Worksheet
    Dim vol_Total As Double
    Dim lastrow As Double
    Dim Summary_Table_Row As Integer
    Dim yearly_change As Double
    Dim yrl_chg_pct As Double
    Dim open_price As Double
    Dim close_price As Double

    
    'Initiailizing the variable vol_total and summary_table_row
    
    
    vol_Total = 0
    Summary_Table_Row = 2
  
    'Initiailizing the loops for each worksheet
  For Each ws In Worksheets
  
    'Calculating last row of each worksheet
    lastrow = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Giving the header to the summary table
    
    ws.Range("I1").Value = "Ticker Name"
    ws.Range("J1").Value = "Total Volume"
    ws.Range("K1").Value = "Close Price"
    ws.Range("L1").Value = "Open Price"
    ws.Range("M1").Value = "Yearly Change"
    ws.Range("N1").Value = "Yearly Change%"
    
    
    'getting first open_price and assigning it summary table row
        open_price = ws.Cells(2, 3).Value
        'ws.Range("L2").Value = open_price
    
    
    'going through each row in each spreadshow begining row number 2 to last row number calculated above
        For i = 2 To lastrow
    
    
     'if values in next row is not equal to current row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

          'Get the ticker name
            ticker_Name = ws.Cells(i, 1).Value
            
            'Get the volume total
            vol_Total = vol_Total + ws.Cells(i, 7).Value
              
            'Ger the closing price
            close_price = ws.Cells(i, 6)
            
            'calculating yearly change
            
            yearly_change = close_price - open_price
            
            'calculating yearly percent change
            yrl_chg_pct = yearly_change * 100 / open_price
            
            On Error Resume Next
            
            'assigning values to the cells
                ws.Range("I" & Summary_Table_Row).Value = ticker_Name
                ws.Range("J" & Summary_Table_Row).Value = vol_Total
                ws.Range("K" & Summary_Table_Row).Value = close_price
                ws.Range("L" & Summary_Table_Row).Value = open_price
                ws.Range("M" & Summary_Table_Row).Value = yearly_change
                ws.Range("N" & Summary_Table_Row).Value = yrl_chg_pct
               
            'assigning color formatting to the cells
            If yearly_change >= 0 Then
            
               ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 4
            
            Else
               ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
                
               'incrementing summary row table
            Summary_Table_Row = Summary_Table_Row + 1
            'getting the open price for next cell
            open_price = ws.Cells(i + 1, 3).Value
            
              'reinitiaizing volume total
            vol_Total = 0

        Else
         'calculating volume total if the ticker is same
            vol_Total = vol_Total + ws.Cells(i, 7).Value

        End If
        'going to the next row
        Next i
        
        'reinitializing summary table
        Summary_Table_Row = 2

 'calculating greatest volume, percent increase and percent decrease
 
    Dim j As Integer
    Dim ticker_Name1, ticker_Name2, ticker_Name3 As String
    Dim lastrow_ticker, high_vol, high_per, low_per As Double
 
    a = 2
    high_vol = ws.Cells(j, 10).Value
    high_per = ws.Cells(j, 14).Value
    low_per = ws.Cells(j, 14).Value
    
    lastrow_ticker = 0
    high_vol = 0
    high_per = 0
    low_per = 0
    
    lastrow_ticker = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ws.Range("Q1").Value = "Summary Type"
    ws.Range("R1").Value = "Ticker Name"
    ws.Range("S1").Value = "Reported Value"
    
       
    For a = 3 To lastrow_ticker
    
        If ws.Cells(a, 10).Value > high_vol Then
            high_vol = ws.Cells(a, 10).Value
            ticker_Name1 = ws.Range("I" & a).Value
          'Else
         End If
         
        If ws.Cells(a, 14).Value > high_per Then
            high_per = ws.Cells(a, 14).Value
            ticker_Name2 = ws.Range("I" & a).Value
         'Else
         End If
         
        If ws.Cells(a, 14).Value < low_per Then
            low_per = ws.Cells(a, 14).Value
            ticker_Name3 = ws.Range("I" & a).Value
         'Else
        
         End If
         
    Next a
    
    ws.Range("Q" & 2).Value = "Greatest Volume"
    ws.Range("R" & 2).Value = ticker_Name1
    ws.Range("S" & 2).Value = high_vol
 
    ws.Range("Q" & 3).Value = "Greatest Increase"
    ws.Range("R" & 3).Value = ticker_Name2
    ws.Range("S" & 3).Value = high_per
  
    ws.Range("Q" & 4).Value = "Greatest Decrease"
    ws.Range("R" & 4).Value = ticker_Name3
    ws.Range("S" & 4).Value = low_per
 
 'Next Worksheet
 Next ws
'End of function
End Sub

