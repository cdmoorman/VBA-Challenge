Attribute VB_Name = "Module1"
Sub alpha_test():

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        
        'Create variable to hold Worksheet Name
            Dim sheetname As String
                    
        'WorksheetName
            sheetname = WS.Name
        
        ' Insert headers to Columns 9 through 12, respectively "Ticker", "Yearly Change", "Percent Change", and "Total Stock Volume"
                  
            WS.Cells(1, 9).Value = "Ticker"
            WS.Cells(1, 10).Value = "Yearly Change"
            WS.Cells(1, 11).Value = "Percent Change"
            WS.Cells(1, 12).Value = "Total Stock Volume"
             
        'Set an initial variable for holding the Stock Ticker symbol
            Dim stock_name As String
    
        'Set an initial variable for holding the total stock volume per ticker
            Dim Total_Volume As Double
            Total_Volume = 0
        
        'Set an initial variable for holding the open price on 1st day of the year
            Dim Open_Price As Long
            Open_Price = WS.Cells(2, 3).Value
            
        
        'Set an initial variable for holding the closing price on last day of the year
            Dim Close_Price As Long
        
        'Set an initial change in price for stock name
            Dim Price_Change As Double
            
         'Set an % change in price for stock name
            Dim Percent_Change As Double
        
        'Create a summary table for each of the variables calculated
            Dim Summary_Table As Integer
            Summary_Table = 2
                    
        ' Determine the Last Row
            lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
            
                   
        'Loop through all of the stock symbols
            For i = 2 To lastrow
        
            'Check if we are still within the same stock ticker
                If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                
                'Set the stock_name
                stock_name = WS.Cells(i, 1).Value
                
                'Enter the stock symbol into the column for "Ticker"
                WS.Range("i" & Summary_Table).Value = stock_name
                
                'Add to the stock volume Total
                Total_Volume = Total_Volume + WS.Cells(i, 7).Value
                
                'Put the stock volume total into Column L
                WS.Range("l" & Summary_Table).Value = Total_Volume
                                                    
                'Set closing price value
                Close_Price = WS.Cells(i, 6).Value
                
                'Determine price change
                Price_Change = Close_Price - Open_Price
                
                'Put the price change total into Column J
                WS.Range("J" & Summary_Table).Value = Price_Change
                
                'Determine price change
                  If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Price_Change / Open_Price
                    
                End If
                
                'Put the price change total into Column J
                WS.Range("K" & Summary_Table).Value = Percent_Change
                WS.Range("K" & Summary_Table).NumberFormat = "0.00%"
                
                'Add a row to the summary table
                Summary_Table = Summary_Table + 1
                
                'Set open price value
                Open_Price = WS.Cells(i + 1, 3).Value
                
                'Reset the stock volume total
                Total_Volume = 0
                
                                
                'If the cell immediately following a row is the same stock ticker...
                Else
                
                    'Add the total stock volume
                    Total_Volume = Total_Volume + WS.Cells(i, 7).Value
                    
                
                End If
                
            Next i
            
            ' Define the Last Row of Summary_Table
                ST_LastRow = WS.Cells(Rows.Count, 10).End(xlUp).Row
            
            ' Set the Cell Colors
                For j = 2 To ST_LastRow
                
                'Set the color of cell to green if price change is greater than or equal to 0
                If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
                
                'Set the color of the cell to red if price change is less than 0
                ElseIf Cells(j, 10).Value < 0 Then
                
                Cells(j, 10).Interior.ColorIndex = 3
                
                End If
        
            Next j
            
        ' Insert column and row headers to Summarize each tab's "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
                  
            WS.Cells(1, 16).Value = "Ticker"
            WS.Cells(1, 17).Value = "Value"
            WS.Cells(2, 15).Value = "Greatest % Increase"
            WS.Cells(3, 15).Value = "Greatest % Decrease"
            WS.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        'Determine last row of stock summary
            SS_lastrow = WS.Cells(Rows.Count, 9).End(xlUp).Row
                       
        'Loop through stock summary table to find stock tickers with the max % increase, greatest % decrease, and greatest total volume
        
            For k = 2 To SS_lastrow
                If Cells(k, 11).Value = Application.WorksheetFunction.max(WS.Range("K2:K" & SS_lastrow)) Then
                    Cells(2, 16).Value = Cells(k, 9).Value
                    Cells(2, 17).Value = Cells(k, 11).Value
                    Cells(2, 17).NumberFormat = "0.00%"
                
                ElseIf Cells(k, 11).Value = Application.WorksheetFunction.min(WS.Range("K2:K" & SS_lastrow)) Then
                    Cells(3, 16).Value = Cells(k, 9).Value
                    Cells(3, 17).Value = Cells(k, 11).Value
                    Cells(3, 17).NumberFormat = "0.00%"
                
                ElseIf Cells(k, 12).Value = Application.WorksheetFunction.max(WS.Range("L2:L" & SS_lastrow)) Then
                    Cells(4, 16).Value = Cells(k, 9).Value
                    Cells(4, 17).Value = Cells(k, 12).Value
                           
            End If
        
        Next k
    
    Next WS
    
    
End Sub
