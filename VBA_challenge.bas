Attribute VB_Name = "Module1"
Sub stockLittle()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

Dim tickerSymbol As String
Dim stockVolume As LongLong
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim lastRow As Long
Dim currentTicker As String
Dim recordTicker As Long
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestTotalVolume As LongLong
Dim lastSubRow As Long
Dim greatestIncreaseRow As Long
Dim greatestDecreaseRow As Long
Dim greatestTotalVolumeRow As Long


ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "YearlyChange"
ws.Range("K1").Value = "PercentChange"
ws.Range("L1").Value = "TotalStockVolume"
ws.Range("N2").Value = "Greatest % increase"
ws.Range("N3").Value = "Greatest % decrease"
ws.Range("N4").Value = "Greatest total volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"


lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

currentTicker = ""
recordTicker = 1
stockVolume = 0
greatestIncrease = 0
greatestDecrease = 0
greatestTotalVolume = 0
greatestIncreaseRow = 0
greatestDecreaseRow = 0
greatestTotalVolumeRow = 0

For i = 2 To lastRow
    If ws.Cells(i, 1) <> currentTicker Then
        'New line for recording
        recordTicker = recordTicker + 1
    
        'Get Ticker
        currentTicker = ws.Cells(i, 1).Value
        'MsgBox ("Row:" & i & " Ticker:" & currentTicker)
        
        'Place Ticker in new ticker column
        ws.Cells(recordTicker, 9).Value = currentTicker
        'Reset intial stock volume
        stockVolume = 0
        
        'get first open price
        openPrice = ws.Cells(i, 3).Value
        
        
    End If
        stockVolume = ws.Cells(i, 7).Value + stockVolume
        ws.Cells(recordTicker, 12).Value = stockVolume
        
        'gets close price every loop
        closePrice = ws.Cells(i, 6).Value
        'calc of yearly change
        yearlyChange = closePrice - openPrice
        'write yearly change into cell
        ws.Cells(recordTicker, 10).Value = yearlyChange
        'calc percent change
        percentChange = (yearlyChange / openPrice)
        'write percent change
        ws.Cells(recordTicker, 11).Value = FormatPercent(percentChange)
    
    If ws.Cells(recordTicker, 10).Value >= 0 Then
        ws.Cells(recordTicker, 10).Interior.ColorIndex = 4
        ws.Cells(recordTicker, 11).Interior.ColorIndex = 4
    Else
        ws.Cells(recordTicker, 10).Interior.ColorIndex = 3
        ws.Cells(recordTicker, 11).Interior.ColorIndex = 3
    End If
        
Next i

lastSubRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

For i = 2 To lastSubRow
    If ws.Cells(i, 11).Value > greatestIncrease Then
        greatestIncrease = ws.Cells(i, 11).Value
        greatestIncreaseRow = i
    ElseIf ws.Cells(i, 11).Value < greatestDecrease Then
        greatestDecrease = ws.Cells(i, 11).Value
        greatestDecreaseRow = i
    End If
    
    If ws.Cells(i, 12).Value > greatestTotalVolume Then
        greatestTotalVolume = ws.Cells(i, 12).Value
        greatestTotalVolumeRow = i
    End If
Next i


ws.Cells(2, 15).Value = ws.Cells(greatestIncreaseRow, 9).Value
ws.Cells(3, 15).Value = ws.Cells(greatestDecreaseRow, 9).Value
ws.Cells(4, 15).Value = ws.Cells(greatestTotalVolumeRow, 9).Value
ws.Cells(2, 16).Value = FormatPercent(greatestIncrease)
ws.Cells(3, 16).Value = FormatPercent(greatestDecrease)
ws.Cells(4, 16).Value = greatestTotalVolume


Next ws
'loop through all sheets
'dont forget conditional formatting (through vba?)

MsgBox ("Done!")



End Sub

