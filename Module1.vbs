Public Sub Stock_Analysis()
  Call turnoff
    'Declarations
    Dim timetaken As Double
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim data As Variant
    Dim lastRow As Long
    Dim numRows As Long
    Dim i As Long
    Dim TickerSymbol As String
    Dim stockData As Object
    Dim currentStock As StockDataClass
    Dim nonBlankCounter As Long
    'Dim debugOutputs As Collection
    Dim outputWs As Worksheet
    Dim infoArray As Variant ' Variable for storing debugOutput data
    'Dim debugOutput As Object ' Variable for each debugOutput array --din'twork
    Dim tickeri As String
    Dim tickerd As String
    Dim tickerg As String
    Dim increasep As Long
    Dim decreasep As Long
    Dim Gvolume As Double
    Dim arr As Variant
    ReDim arr(9, 3)
    Dim outputA As Variant
    Dim outputB As Variant
    Dim outputC As Variant
    ' Initialize a collection to store data arrays for each worksheet-- din't work
    'Set debugOutputs = New Collection
    timetaken = MicroTimer
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Check if the worksheet contains any tables
        If ws.ListObjects.Count > 0 Then
            ' Loop through all tables in the worksheet
            'Sort the table
            'tbl.sort
            
            For Each tbl In ws.ListObjects
                ' Read data from the table into arrays
                data = tbl.DataBodyRange.Value
                lastRow = tbl.ListRows.Count
                numRows = lastRow

                ' Initialize variables for the current worksheet
                Set stockData = CreateObject("Scripting.Dictionary")
                nonBlankCounter = 0

                ' Initialize an array to store data for this table
                Dim debugOutput() As Variant
                ReDim debugOutput(1 To numRows, 1 To 4)

                ' Loop through the array and calculate information for each ticker symbol
                For i = 1 To numRows - 1
                    TickerSymbol = data(i, 1)
                    ' If the ticker symbol is not in the dictionary, create a new StockData object
                    If Not stockData.Exists(TickerSymbol) Then
                        Dim newStock As New StockDataClass
                        newStock.TickerSymbol = TickerSymbol
                        newStock.openPrice = data(i, 3) 'Assign the first day open price
                        newStock.totalVolume = 0 ' Initialize totalVolume for the new ticker symbol
                        stockData.Add TickerSymbol, newStock
                    End If

                    ' Update the StockData object with the close price and volume
                    stockData(TickerSymbol).closePrice = data(i, 6)
                    stockData(TickerSymbol).AccumulateVolume CDbl(data(i, 7))

                    ' If a new ticker symbol starts or it's the last row, calculate yearly change and percentage change
                    If i = numRows - 1 Or TickerSymbol <> data(i + 1, 1) Then
                        Set currentStock = stockData(TickerSymbol)

                        ' Calculate the percentage change for the current StockData object
                        currentStock.CalculatePercentageChange

                        ' Store the information in the array for later writing
                        nonBlankCounter = nonBlankCounter + 1
                        debugOutput(nonBlankCounter, 1) = currentStock.TickerSymbol
                        debugOutput(nonBlankCounter, 2) = Round(currentStock.closePrice - currentStock.openPrice, 2)
                        debugOutput(nonBlankCounter, 3) = Round(currentStock.percentageChange, 2) & "%"
                        debugOutput(nonBlankCounter, 4) = currentStock.totalVolume
                        
                        If nonBlackCounter = 1 Then
                            tickeri = currentStock.TickerSymbol
                            tickerd = currentStock.TickerSymbol
                            tickerg = currentStock.TickerSymbol
                            increasep = Round(currentStock.percentageChange, 2)
                            decreasep = Round(currentStock.percentageChange, 2)
                            Gvolume = currentStock.totalVolume
                        ElseIf increasep < currentStock.percentageChange Or decreasep > currentStock.percentageChange Or Gvolume < currentStock.totalVolume Then
                            If increasep < currentStock.percentageChange Then
                            increasep = currentStock.percentageChange
                            tickeri = currentStock.TickerSymbol
                            ElseIf decreasep > currentStock.percentageChange Then
                             decreasep = currentStock.percentageChange
                             tickerd = currentStock.TickerSymbol
                            ElseIf Gvolume < currentStock.totalVolume Then
                             tickerg = currentStock.TickerSymbol
                             Gvolume = currentStock.totalVolume
                            End If
                        
                        End If
                        
                        ' Remove the ticker symbol from the dictionary
                        stockData.Remove TickerSymbol
                    End If
                Next i
            Next tbl
        End If
       If ws.Name = "2018" Then
          
          outputA = debugOutput
         
          arr(1, 1) = "Greatest % Increase"
          arr(1, 2) = tickeri
          arr(1, 3) = increasep
          arr(2, 1) = "Greatest % Decrease"
          arr(2, 2) = tickerd
          arr(2, 3) = decreasep
          arr(3, 1) = "Total Stock Volume"
          arr(3, 2) = tickerg
          arr(3, 3) = Gvolume
        ElseIf ws.Name = "2019" Then
          
          outputB = debugOutput
          arr(4, 1) = "Greatest % Increase"
          arr(4, 2) = tickeri
          arr(4, 3) = increasep
          arr(5, 1) = "Greatest % Decrease"
          arr(5, 2) = tickerd
          arr(5, 3) = decreasep
          arr(6, 1) = "Total Stock Volume"
          arr(6, 2) = tickerg
          arr(6, 3) = Gvolume
        ElseIf ws.Name = "2020" Then
         
          outputC = debugOutput
          arr(7, 1) = "Greatest % Increase"
          arr(7, 2) = tickeri
          arr(7, 3) = increasep
          arr(8, 1) = "Greatest % Decrease"
          arr(8, 2) = tickerd
          arr(8, 3) = decreasep
          arr(9, 1) = "Total Stock Volume"
          arr(9, 2) = tickerg
          arr(9, 3) = Gvolume
 
    
        End If
         tickeri = " "
         tickerd = " "
         tickerg = " "
         increasep = 0
         decreasep = 0
         Gvolume = 0
    Next ws
    
   'Writing on to each work sheet.
   For Each ws In ThisWorkbook.Sheets
       
        If ws.Name = "2018" Then
             Set outputWs = ThisWorkbook.Sheets("2018")
             Call write_data(ws, outputA, 1, arr)
        ElseIf ws.Name = "2019" Then
             Set outputWs = ThisWorkbook.Sheets("2019")
             Call write_data(ws, outputB, 4, arr)
        ElseIf ws.Name = "2020" Then
             Set outputWs = ThisWorkbook.Sheets("2020")
             Call write_data(ws, outputC, 7, arr)
        End If
   

   Next ws
   Call turnon
  Debug.Print "Time Taken is: " & (MicroTimer - timetaken) * 1000
End Sub

Sub write_data(ws As Worksheet, output As Variant, i As Integer, sarr As Variant)
  
  'Declarations
  'Dim nextOutputRow As Long
   Dim lrow As Long
   Dim rowj As Long
  
   'Find number rows in array.
   lrow = UBound(output) - LBound(output) + 1
'clear the contents
   ws.Range("I:Q").CurrentRegion.ClearContents
      'Write the Header
   ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
   'Write the summary by ticker ,yearly ,percent change and total volume
   ws.Range("I2:L" & lrow).Value = output
   'Color coding for yearlychange.
   For rowj = 1 To lrow
     If output(rowj, 2) > 0 Then
        ws.Cells(rowj + 1, 10).Interior.Color = RGB(0, 255, 0)
     Else
        ws.Cells(rowj + 1, 10).Interior.Color = RGB(255, 0, 0)
     End If
    Next rowj
     
  ' Writing the Greatest increase or descrease in percent change and volume
   
   ws.Range("P1:Q1") = Array("Ticker", "Value")
   
   Dim k, j, l, c, wr As Integer
   wc = 15
   c = 1
   wr = 2
   For k = i To i + 2
     For c = 1 To 3
        ws.Cells(wr, wc) = sarr(k, c)
        wc = wc + 1
         
    Next c
         wc = 15
         wr = wr + 1
   Next k
   
End Sub

Sub turnoff()
 Application.Calculation = xlCalculationManual
 Application.ScreenUpdating = False
 Application.EnableEvents = False
End Sub

Sub turnon()

 Application.Calculation = xlCalculationAutomatic
 Application.ScreenUpdating = True
 Application.EnableEvents = True

End Sub

