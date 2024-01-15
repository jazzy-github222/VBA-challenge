Attribute VB_Name = "Module3"
Sub Yearly_Stock_Data()
    
    ' VARIABLES FOR PART 1
    ' --------------------------------
        ' Declare variables for Table 1
        Dim last_row As Long
        Dim ws As Worksheet
        Dim ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percentage_change As Double
        Dim total_stock_volume As Double
        Dim table_1_rows As Long
  
     ' HEADINGS FOR PART 1
    ' --------------------------------
   ' Loop through all worksheets
    For Each ws In Worksheets
       ' Declare where the values will appear
        ' Remember to use "NumberFormat = "0.00%" for percentages
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
                ' Make the headings bold
                ws.Range("I1:P1").Font.Bold = True
        
        ' Find the last row with data in column A
        last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
  
   ' VARIABLES FOR PART 2
    ' --------------------------------
        '  Declare variables for Table 2
         Dim max_increase As Double
         Dim max_decrease As Double
         Dim max_total_volume As Double
         Dim MaxTickerInc As String
         Dim MaxTickerDec As String
         Dim MaxTicker_TotalVolume As String
         
        ' KEY variables for Table 1
        table_1_rows = 2
        ticker = ws.Cells(2, 1).Value
        open_price = ws.Cells(2, 3).Value
        total_stock_volume = 0
        
        
    ' BEGIN THE HERE CODE BY:
    ' -----------------------------------------------------
        ' Looping through the data accross worksheets
        For i = 2 To last_row
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    ' Calculate the yearly changes by:
                    '  Declaring the closing price and...
                    close_price = ws.Cells(i, 6).Value
                    yearly_change = close_price - open_price
                    
                            ' ...calculating the yearly change, and percentage change
                            If open_price <> 0 Then
                                percentage_change = (yearly_change / open_price)
                            Else
                                percentage_change = 0
                            End If
                
                    ' Determine where the results will go in Table 1
                      ws.Cells(table_1_rows, 9).Value = ticker
                      ws.Cells(table_1_rows, 10).Value = yearly_change
                      
                    ' Use number formatting to easily change into a percentage
                      ws.Cells(table_1_rows, 11).Value = percentage_change
                      ws.Cells(table_1_rows, 11).NumberFormat = "0.00%"
                      
                      ws.Cells(table_1_rows, 12).Value = total_stock_volume
        
                              ' Check for the "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume"
                            If percentage_change > max_increase Then
                                  max_increase = percentage_change
                                  MaxTickerInc = ticker
                
                                ElseIf percentage_change < max_decrease Then
                                     max_decrease = percentage_change
                                     MaxTickerDec = ticker
                                
                                ElseIf total_stock_volume > max_total_volume Then
                                     max_total_volume = total_stock_volume
                                     MaxTicker_TotalVolume = ticker
                                     
                            End If
                      
                             ' Conditional formatting for positive change in green and negative change in red
                            If yearly_change > 0 Then
                                ws.Cells(table_1_rows, 10).Interior.ColorIndex = 4 ' Green
                                
                                ElseIf yearly_change < 0 Then
                                    ws.Cells(table_1_rows, 10).Interior.ColorIndex = 3 ' Red
                                    
                            End If
                
                    ' Move to the next line in the summary table
                    table_1_rows = table_1_rows + 1
                    
                    ' Reset variables for the new ticker
                    ticker = ws.Cells(i + 1, 1).Value
                    open_price = ws.Cells(i + 1, 3).Value
                    total_stock_volume = 0
             
                Else
                
                ' Accumulate total stock volume for the current ticker
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
         ' HEADINGS FOR PART 2
        ' --------------------------------
        ' Write the "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume" information in Table 2
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Range("N2:N4").Font.Bold = True
    
         ' VARIABLE LOCATIONS FOR PART 2
        ' --------------------------------
        ws.Cells(2, 15).Value = MaxTickerInc
        ws.Cells(3, 15).Value = MaxTickerDec
        ws.Cells(4, 15).Value = MaxTicker_TotalVolume
    
        ws.Cells(2, 16).Value = max_increase
        ws.Cells(3, 16).Value = max_decrease
        ws.Cells(4, 16).Value = max_total_volume
        
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = max_total_volume
    
   
   Next ws
    
End Sub

