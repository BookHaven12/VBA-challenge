# VBA-challenge
Places I got help for my code:
1.  I used the Xpert Learning Assistant for the following:
   If i = lastrow Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then,

   formatting as a percentage:  ws.Range("Q3").NumberFormat = "0.00%" ' Format as percentage
   
   No color change for a value of zero: ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = xlNone

3. I asked for help from Paul during office hours about Dim StockTotal As LongLong, and for help with why my calculation for quarterly change. It was calculating incorrectly

4. To loop through multiple workbooks:
https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
