Sub forLoop()
    
    ' Create variables for the various calculations
    Dim ticker, biggestWinner, biggestLoser, biggestMover As String
    Dim openPrice, closePrice, yearlyPriceChange, yearlyPctChange, greatestInc, greatestDec As Double
    Dim totalStockVol As LongLong
    Dim greatestTotalVol As LongLong
    Dim lastRow As Long
    Dim tickerCount As Integer
    Dim cond1 As FormatCondition, cond2 As FormatCondition
    
    ' Set values for variables
    tickerCount = 1
    ticker = ""
    totalStockVol = 0
    greatestTotalVol = 0
    greatestInc = 0
    greatestDec = 0
    
    ' Formating
    Range("K:K").NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = Number
    
    Set rg = Range("J2", Range("J2").End(xlDown))        
    Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreaterEqual, "0")
    Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "0")

    With cond1
    .Interior.Color = vbGreen
    End With
    
    With cond2
    .Interior.Color = vbRed
    End With
    
    ' Set up headers for the columns in the report
    Cells(1, 9).Value = "Ticker"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 17).Value = "Value"

    ' Figure out how many rows there are, to determine the iteration range below
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through the tickers
    For i = 2 To lastRow
        ' This first block is searches for the first entry of a new ticker. It grabs the Open Price, 
        ' and sets the tickerCount which will determine which row the summary data gets recorded in. 
        If Cells(i, 1).Value <> ticker Then
            totalStockVol = Cells(i, 7).Value
            tickerCount = tickerCount + 1
            ticker = Cells(i, 1).Value
            Cells(tickerCount, 9).Value = ticker
            openPrice = Cells(i, 3).Value

        ' This else statement will keep adding stock volume for all the rows after the first one. 
        Else
            totalStockVol = totalStockVol + Cells(i, 7).Value
        
        End If
        
        ' This next block is intended to identify the LAST row of a given ticker symbol. When it detects
        ' a final row, it will grab the Close Price for that day, and it will then calculate and record the
        ' data points that the analysis asks for like Price Change and Percent Change.

        If Cells(i + 1, 1).Value <> ticker Then
            closePrice = Cells(i, 6).Value
            yearlyPriceChange = closePrice - openPrice
            If openPrice <> 0 Then
                yearlyPctChange = yearlyPriceChange / openPrice
            Else: yearlyPctChange = 0
            End If
            Cells(tickerCount, 10).Value = yearlyPriceChange
            Cells(tickerCount, 11).Value = yearlyPctChange
            Cells(tickerCount, 12).Value = totalStockVol
        
        End If
    Next i
    
    ' This next blog searches for the biggest movers. 
    For i = 2 To tickerCount

        'This sub-block looks for the ticker with the biggest gains. 
        If Cells(i, 11).Value > greatestInc Then
            greatestInc = Cells(i, 11).Value
            biggestWinner = Cells(i, 9).Value
        End If
        
        'This sub-block looks for the ticker with the biggest losses. 
        If Cells(i, 11).Value < greatestDec Then
            greatestDec = Cells(i, 11).Value
            biggestLoser = Cells(i, 9).Value
        End If
        
        ''This sub-block looks for the ticker with the biggest transaction volume. 
        If Cells(i, 12).Value > greatestTotalVol Then
            greatestTotalVol = Cells(i, 12).Value
            biggestMover = Cells(i, 9).Value
        
       End If
    Next i

    ' This last section simply prints the results of the previous block on to the sheet.     
    Cells(2, 15).Value = "Greatest % increase"
    Cells(2, 16).Value = biggestWinner
    Cells(2, 17).Value = greatestInc
    
    Cells(3, 15).Value = "Greatest % decrease"
    Cells(3, 16).Value = biggestLoser
    Cells(3, 17).Value = greatestDec
    
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 16).Value = biggestMover
    Cells(4, 17).Value = greatestTotalVol
    
 
End Sub






