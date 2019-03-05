Attribute VB_Name = "Module1"
Sub Alphabtical_Test()

' Easy
' Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
' You will also need to display the ticker symbol to coincide with the total volume.

' Moderate
' Create a script that will loop through all the stocks and take the following info.
' Yearly change from what the stock opened the year at to what the closing price was.
' The percent change from the what it opened the year at to what it closed.
' The total Volume of the stock
' Ticker Symbol
' You should also have conditional formatting that will highlight positive change in green and negative change in red.

' Hard
' Your solution will include everything from the moderate challenge.
' Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".


' Define Variables
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim Ticker_Row As Long
Dim NextTickerOpenPrice As Double
Dim CurrentTickerOpenPrice As Double
Dim ClosePrice As Double
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotalVolume As Long
Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestVolumeTicker As String

            

    ' Loop thru all sheets
    For Each ws In Worksheets
    
    ' Initialize the start
    TotalStockVolume = 0
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0

    ' Set Ticker Arrays for Ticker summary table
    TickerSum = 2
    
    ' Set the Ticker opening price
    NextTickerOpenPrice = ws.Cells(2, 3).Value
    
    ' Define the LastRow
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Add Header
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    
    
        For i = 2 To LastRow
                   
            ' Check we are still within the same Ticker, if it is not..
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
            ' Bring the Ticker name and total to the Summary
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & TickerSum).Value = Ticker
            
                
            ' Calculate Yearly Change Value
            CurrentTickerOpenPrice = NextTickerOpenPrice
            CurrentTickerClosePrice = ws.Cells(i, 6).Value
            YearlyChange = CurrentTickerClosePrice - CurrentTickerOpenPrice
            ws.Range("J" & TickerSum).Value = YearlyChange
            
                ' Set Conditional Formatting for Yearly Change Value
                If YearlyChange >= 0 Then
                ws.Range("J" & TickerSum).Interior.ColorIndex = 4 ' Green
                Else
                ws.Range("J" & TickerSum).Interior.ColorIndex = 3 ' Red
                End If
                
                '  Calculate Yearly Change Percentage
                If CurrentTickerOpenPrice = 0 Then
                ws.Cells(TickerSum, 10) = "NA"
                Else
                PercentChange = YearlyChange / CurrentTickerOpenPrice
                ws.Range("K" & TickerSum).Value = PercentChange

                End If
                
          
            ' Calculate Total Stock Volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            ws.Range("L" & TickerSum).Value = TotalStockVolume
            TotalStockVolume = 0
            
            ' Add one to the summary table row
            TickerSum = TickerSum + 1
            
            ' Grab the next ticker opening price
            NextTickerOpenPrice = ws.Cells(i + 1, 3).Value
            

            
            Else
            
            
            ' Calculate Total Stock Volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                

            End If
            
            
             
        Next i
        
            ' Calculate Gretest Increase %
            GreatestIncrease = WorksheetFunction.Max(ws.Range("K:K"))
            ws.Range("P2") = GreatestIncrease
            ws.Range("P2") = Format(GreatestIncrease, "0.00%")
                
            ' Calculate Gretest Decrease %
            GreatestDecrease = WorksheetFunction.Min(ws.Range("K:K"))
            ws.Range("P3") = GreatestDecrease
            ws.Range("P3") = Format(GreatestDecrease, "0.00%")

            ' Calculate Gretest Volume
            GreatestVolume = WorksheetFunction.Max(ws.Range("L:L"))
            ws.Range("P4") = GreatestVolume
         
            ' Show Greatest Increase Ticker
            GreatestIncreaseTicker = ws.Range("I" & WorksheetFunction.Match(GreatestIncrease, ws.Range("K:K"), 0))
            ws.Range("O2") = GreatestIncreaseTicker
            
            ' Show Greatest Decrease Ticker
            GreatestDecreaseTicker = ws.Range("I" & WorksheetFunction.Match(GreatestDecrease, ws.Range("K:K"), 0))
            ws.Range("O3") = GreatestDecreaseTicker
            
            ' Show Greatest Volume Ticker
            GreatestDecreaseTicker = ws.Range("I" & WorksheetFunction.Match(GreatestVolume, ws.Range("L:L"), 0))
            ws.Range("O4") = GreatestDecreaseTicker
            
            ' Set width and Format
            ws.Columns("A:Q").AutoFit
            
    Next ws
    
MsgBox("Complete")

End Sub

