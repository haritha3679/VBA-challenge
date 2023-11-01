'Class Module: stockData

Public TickerSymbol As String
Public openPrice As Double
Public closePrice As Double
Public totalVolume As Double
Public percentageChange As Double

Public Sub CalculatePercentageChange()
    If openPrice <> 0 Then
        percentageChange = ((closePrice - openPrice) / openPrice) * 100
    Else
        percentageChange = 0
    End If
End Sub

Public Sub AccumulateVolume(volume As Double)
    totalVolume = totalVolume + volume
End Sub



