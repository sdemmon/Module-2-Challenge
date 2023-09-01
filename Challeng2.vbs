Sub Stocks()

    'Create vrables to hold values
    Dim LR As Long
    Dim ticker As String
    Dim ClosePrice As Double
    Dim OpenPrice As Double
    Dim YearlyChange As Double
    Dim StockVolume As Double
    Dim ws As Worksheet
    
    'loop through all sheets
    For Each ws In Worksheets
    
    'Set initial stock volume and row count
    StockVolume = 0
    RowCount = 2

    'determine last row
    LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Heading Summary
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "YearlyChange"
ws.Cells(1, 11).Value = "PercentageChange"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

'set initial open price
OpenPrice = ws.Cells(2, 3).Value

For i = 2 To LR

'check if we are still within the same ticker
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'set ticker name
    ticker = ws.Cells(i, 1).Value
    ws.Cells(RowCount, 9).Value = ticker
    'add total volume
    StockVolume = StockVolume + ws.Cells(i, 7).Value
    
    'set close price
    ClosePrice = ws.Cells(i, 6).Value
    'add yearly change
    YearlyChange = (OpenPrice - ClosePrice)
    'add percentage change
    PercentChange = (YearlyChange / OpenPrice)
    'insert values to summary table
    ws.Cells(RowCount, 10).Value = YearlyChange
    ws.Cells(RowCount, 11).Value = PercentChange
    ws.Cells(RowCount, 12).Value = StockVolume
    ws.Cells(RowCount, 11).NumberFormat = "0.00%"
    'set open price for next loop
    OpenPrice = ws.Cells(i + 1, 3).Value

    'format
    If YearlyChange > 0 Then
        ws.Cells(RowCount, 10).Interior.Color = RGB(0, 255, 0)
    ElseIf YearlyChange < 0 Then
        ws.Cells(RowCount, 10).Interior.Color = RGB(255, 0, 0)
    End If
    
    'reset counts
    RowCount = RowCount + 1
    YearlyChange = 0
    StockVolume = 0
    
Else
    'add volume to total volume within a ticker
    StockVolume = StockVolume + ws.Cells(i, 7).Value

End If

Next i

'create variables to hold values
Dim minValue As Double
Dim maxValue As Double
Dim greatVol As Double

'determine last row for yearly change
LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Look through each rows to find the associated value and its associate ticker
For Z = 2 To LastRow
    If ws.Cells(Z, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) Then
        maxValue = ws.Cells(Z, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(Z, 11).Value
            ws.Cells(2, 16).NumberFormat = "0.00%"
        ticker = ws.Cells(Z, 9).Value
            ws.Cells(2, 15).Value = ticker
    ElseIf ws.Cells(Z, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) Then
        minValue = ws.Cells(Z, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(Z, 11).Value
            ws.Cells(3, 16).NumberFormat = "0.00%"
        ticker = Cells(Z, 9).Value
            ws.Cells(3, 15).Value = ticker
    ElseIf ws.Cells(Z, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow)) Then
        greatVol = ws.Cells(Z, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(Z, 12).Value
        ticker = ws.Cells(Z, 9).Value
            ws.Cells(4, 15).Value = ticker
            End If
        Next Z

    Next ws

End Sub
