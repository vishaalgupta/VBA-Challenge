Attribute VB_Name = "Module1"
Sub StockInfo()
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Dim counter As Integer
Dim openPrice As Double
Dim closePrice As Double
Dim totalVolume As LongLong
counter = 1
openPrice = Cells(2, 3).Value
totalVolume = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim lastrow2 As Integer
lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To lastrow
    totalVolume = totalVolume + Cells(i, 7).Value
    If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
        counter = counter + 1
        Cells(counter, 9).Value = Cells(i, 1).Value
        closePrice = Cells(i, 6).Value
        Cells(counter, 10).Value = closePrice - openPrice
        Cells(counter, 11).Value = (closePrice - openPrice) / openPrice
        Range("K2:K" & lastrow2).NumberFormat = "0.00%"
        Cells(counter, 12).Value = totalVolume
        totalVolume = 0
        openPrice = Cells(i + 1, 3).Value
    End If
Next i
For j = 2 To lastrow2
    If Cells(j, 10).Value > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
    ElseIf Cells(j, 10).Value < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
    End If
Next j
Dim increase As Double
Dim increaseTicker As String
Dim increasePercent As Double
Dim decrease As Double
Dim decreaseTicker As String
Dim decreasePercent As Double
Dim maxVol As LongLong
Dim maxVolTicker As String
increase = 0
decrease = 0
maxVol = 0
For k = 2 To lastrow2
    If Cells(k, 11).Value > increase Then
        increase = Cells(k, 11).Value
        increasePercent = Cells(k, 11).Value
        increaseTicker = Cells(k, 9).Value
    End If
    If Cells(k, 11).Value < decrease Then
        decrease = Cells(k, 11).Value
        decreasePercent = Cells(k, 11).Value
        decreaseTicker = Cells(k, 9).Value
    End If
    If Cells(k, 12).Value > maxVol Then
        maxVol = Cells(k, 12).Value
        maxVolTicker = Cells(k, 9).Value
    End If
Next k
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P2").Value = increaseTicker
Range("Q2").Value = increasePercent
Range("P3").Value = decreaseTicker
Range("Q3").Value = decreasePercent
Range("P4").Value = maxVolTicker
Range("Q4").Value = maxVol
Range("Q2:Q3").NumberFormat = "0.00%"
End Sub
