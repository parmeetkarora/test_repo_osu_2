Attribute VB_Name = "Module1"
Sub alphabetical_testingPKA()

    'loop through all sheets
For Each ws In Worksheets

'set the header as the requasted
ws.Cells(1, "i").Value = "Ticker"
ws.Cells(1, "j").Value = "Quarterly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"
ws.Cells(2, "O").Value = "Greatest % increase"
ws.Cells(3, "O").Value = "Greatest % decrease"
ws.Cells(4, "O").Value = "Greatest total volume"
ws.Cells(1, "P").Value = "Ticker"
ws.Cells(1, "Q").Value = "Value"

'identify all variables in part I
Dim current_ticker As String
Dim output_index As Double
output_index = 2
Dim QuarterlyChange As Double
Dim OpeningPrice As Double
OpeningPrice = ws.Cells(2, 3).Value
Dim ClosingPrice As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double

'loop for part I by determining the last row
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastRow

'comparing cells
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'introduce the ticker name
current_ticker = ws.Cells(i, 1).Value

'calculating the changes over a year for each tickers
ClosingPrice = ws.Cells(i, 6).Value
QuarterlyChange = (ClosingPrice - OpeningPrice)

'calculating the change percent for each tickers by using conditions
If OpeningPrice = 0 Then
PercentChange = 0
Else
PercentChange = (QuarterlyChange / OpeningPrice)
End If

'finding the total stock volume for each tickers
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7)

'printing all result in identified cells
ws.Cells(output_index, 9).Value = current_ticker
ws.Cells(output_index, 10).Value = QuarterlyChange
ws.Cells(output_index, 11).Value = PercentChange
ws.Cells(output_index, 11).NumberFormat = "0.00%"
ws.Cells(output_index, 12).Value = TotalStockVolume

'Add a new row
output_index = output_index + 1

'reset the total stock volume
TotalStockVolume = 0

'add new opening price
OpeningPrice = ws.Cells(i + 1, 3).Value
Else
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7)
End If

'conditional formatting for Quarterly change
If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(i, 10).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
Else
ws.Cells(i, 10).Interior.ColorIndex = 0
End If
Next i

'finding the maximum of percent changing and its similar ticker
max_change = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
max_ticker = WorksheetFunction.Match(max_change, ws.Range("K2:K" & lastRow), 0)
ws.Cells(2, 17).Value = max_change
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(2, 16).Value = ws.Cells(max_ticker + 1, 9)

'finding the minimum of percent changing and its similar ticker
min_change = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
min_ticker = WorksheetFunction.Match(min_change, ws.Range("K2:K" & lastRow), 0)
ws.Cells(3, 17).Value = min_change
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = ws.Cells(min_ticker + 1, 9)

'finding the maximum of total stock volume and its similar ticker
max_volume = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
max_total_ticker = WorksheetFunction.Match(max_volume, ws.Range("L2:L" & lastRow), 0)
ws.Cells(4, 17).Value = max_volume
ws.Cells(4, 16).Value = ws.Cells(max_total_ticker + 1, 9)

Next ws

End Sub


