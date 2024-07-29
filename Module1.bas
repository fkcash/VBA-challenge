Attribute VB_Name = "Module1"
Sub QuarterLoops()
Dim lastRow As Long

'Find the last nonblank cell in column A
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
MsgBox (lastRow)
End Sub
Sub StockAnalysis()
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, summaryRow As Long
    Dim tickerSymbol As String, openPrice As Double, closePrice As Double
    Dim quarterlyChange As Double, percentChange As Double, totalVolume As Double
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
    Dim increaseTickerSymbol As String, decreaseTickerSymbol As String, volumeTickerSymbol As String
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        ' Set up summary headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ' Loop through all rows
        For i = 2 To lastRow
            ' Check if it's a new ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Set new ticker and opening price
                tickerSymbol = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
            ' Add to total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            ' Check if it's the last row for the current ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Set closing price
                closePrice = ws.Cells(i, 6).Value
                ' Calculate changes
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = quarterlyChange / openPrice
                Else
                    percentChange = 0
                End If
                ' Output results
                ws.Cells(summaryRow, 9).Value = tickerSymbol
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                ' Conditional formatting
                If quarterlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                ' Check for greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    increaseTickerSymbol = tickerSymbol
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decreaseTickerSymbol = tickerSymbol
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTickerSymbol = tickerSymbol
                End If
                ' Move to next summary row
                summaryRow = summaryRow + 1
            End If
        Next i
        ' Output greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = increaseTickerSymbol
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 16).Value = decreaseTickerSymbol
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 16).Value = volumeTickerSymbol
        ws.Cells(4, 17).Value = greatestVolume
        ' Format cells
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0"
        ' Autofit columns
        ws.Columns("I:Q").AutoFit
    Next ws
    MsgBox "Analysis complete!"
End Sub
