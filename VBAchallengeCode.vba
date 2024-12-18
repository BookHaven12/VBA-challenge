Attribute VB_Name = "Module1"
Sub Mult_year_stock_data()

    ' Declare ws as a worksheet object variable.
    Dim ws As Worksheet

    'Create a variables for the data
    Dim i As Long
    Dim Ticker As String
    Dim Qchange As Double
    Dim PercentChange As Double
    Dim StockTotal As LongLong
    
    'variables for tracking first open and last close prices
    Dim FirstOpen As Double
    Dim LastClose As Double
    
    'create variables for Greatest increase, decrease and total volume
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVol As LongLong
    Dim MaxTicker As String
    Dim IncreaseTicker As String
    Dim DecreaseTicker As String
    Dim TotalVolTicker As String
    
    'Loop through all of the worksheets in the active workbook.
    For Each ws In Worksheets
        
        'initialize tracking variables
        Qchange = 0
        StockTotal = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVol = 0
        
        'keep track of location for each ticker and changes
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        'calculate last row before the loop
        Dim lastrow As Long
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
       ' Set headers for summary table & Greatest inc,decrease and vol
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Reset first open price for the first row
        FirstOpen = ws.Cells(2, 3).Value
        
        'Loopthrough rows in the column
        For i = 2 To lastrow
        
            'check if I reached the last row of the ticker group
            If i = lastrow Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             
               'Set the Ticker Name and LastClose Price
                Ticker = ws.Cells(i, 1).Value
                 LastClose = ws.Cells(i, 6).Value
                
                 'Calculate the quarterly change for the ticker
                 Qchange = LastClose - FirstOpen
                    
                 'Calculate Total Stock Volume for the ticker
                 StockTotal = StockTotal + ws.Cells(i, 7).Value
                
                'Calculate the percentage change for the ticker
                      If FirstOpen <> 0 Then
                    
                        'Prevent division by zero
                         PercentChange = ((LastClose - FirstOpen) / FirstOpen)
                        
                      Else
                          PercentChange = 0
                        
                      End If
                    
                'print data in summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = Qchange
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%" ' Display as percentage
                ws.Range("L" & Summary_Table_Row).Value = StockTotal
        
                   'color formatting for Qchange
                        If Qchange > 0 Then
                           'color green
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                         ElseIf Qchange < 0 Then
                            'color red
                               ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                         Else
                               'no color if the value=0
                               ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = xlNone
                         End If
                         
                'Find Greatest % increase, decrease and total vol
                 If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    IncreaseTicker = Ticker
                End If
                
                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    DecreaseTicker = Ticker
                End If
                
                If StockTotal > GreatestTotalVol Then
                    GreatestTotalVol = StockTotal
                    TotalVolTicker = Ticker
                End If
                
                
                'Increment Summary Table (Add  one to the summary table row)
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset valuesfor FirstOpen, Q Change, and Total Stock Vol
                If i < lastrow Then
                    FirstOpen = ws.Cells(i + 1, 3).Value
                Qchange = 0
                StockTotal = 0
                End If
                
                'if the cell immediately following a row is the same ticker...
            Else
                    'if the cell immediately following a row is the same ticker
                    'add to the total stock volume
                     StockTotal = StockTotal + ws.Cells(i, 7).Value
              End If
            
        Next i
       
       ' Print results for greatest increase, decrease, and total volume
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = IncreaseTicker
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q2").NumberFormat = "0.00%" ' Format as percentage

        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = DecreaseTicker
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q3").NumberFormat = "0.00%" ' Format as percentage

        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = TotalVolTicker
        ws.Range("Q4").Value = GreatestTotalVol

       
   Next ws
End Sub
     
     
            
            
            
    
    
    
    
    
    
    
    
    





