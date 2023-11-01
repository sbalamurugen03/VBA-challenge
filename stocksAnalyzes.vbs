Sub stocksAnalyzes()

 

    ' Loop through all sheets one by one and run the logic

    For Each ws In Worksheets

   

        ' Create a Variable to Hold Open price, Close Price, Total stock volume and ticker Symbol

        Dim OpenPrice As Double

        Dim closePrice As Double

        Dim TotalStockVol As Double

        Dim tickerSymbol As String

       

        Dim column As Integer

        column = 1

        OpenPrice = 0

       

        ' Determine the Last Row

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

       

        ' Assign the header name for the new columns

        ws.Range("I" & 1).Value = "Ticker"

        ws.Range("J" & 1).Value = "Yearly Change"

        ws.Range("K" & 1).Value = "Percent Change"

        ws.Range("L" & 1).Value = "Total Stock Volume"

       

        ' Iterates data (Row vise), calculate and assign the calculated value to the new columns

        For i = 2 To LastRow

       

            ' Searches for when the value of the next cell is different than that of the current cell

            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then

               

                ' Get the ticker symbol, close price and total stock volume, assign to the variables

                tickerSymbol = ws.Cells(i, 1).Value

                closePrice = ws.Cells(i, 6).Value

                TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value

               

                ' Determine next row

                LastRowTicker = ws.Cells(Rows.Count, "I").End(xlUp).Row + 1

 

                ' Assign the calculated value to the new columns

                ws.Range("I" & LastRowTicker).Value = tickerSymbol

                ws.Range("J" & LastRowTicker).Value = closePrice - OpenPrice

                ws.Range("K" & LastRowTicker).Value = ((closePrice - OpenPrice) / OpenPrice)

                ws.Range("L" & LastRowTicker).Value = TotalStockVol

               

                ' Reset the variables

                OpenPrice = 0

                TotalStockVol = 0

                closePrice = 0

               

                ' Assign the cell color based on conditional formatting

                If ws.Cells(LastRowTicker, 10).Value >= 0 Then

                    ' Color the positive value to grade green

                    ws.Cells(LastRowTicker, 10).Interior.ColorIndex = 4

                Else

                   ' Color the negative value to grade red

                    ws.Cells(LastRowTicker, 10).Interior.ColorIndex = 3

 

                End If

               

                'Format the column "Percentage change" to percentage type

                With ws.Cells(LastRowTicker, 11)

                    .Style = "Percent"

                    .NumberFormat = "0.00%"

                End With

               

                'ws.Range("K" & LastRowTicker).Style = "Percent"

 

            Else

           

                'Assign the open price

                If OpenPrice = 0 Then

                    OpenPrice = ws.Cells(i, 3).Value

                End If

                

                ' Calculate the stock volume

                TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value

               

            End If

 

        Next i

       

        ' Determine the Last Row

        LastRowTick = ws.Cells(Rows.Count, "I").End(xlUp).Row

       

        ' create variables to hold the values

        Dim GreatestInc As Double

        Dim GreatestDec As Double

        Dim GreatestTotalStockVol As Double

        Dim GreatestIncTick As String

        Dim GreatestDecTick As String

        Dim GreatestTotalStockVolTick As String

       

        ' Assign the values to the variables

        GreatestInc = 0

        GreatestDec = 0

        GreatestTotalStockVol = 0

      

       ' Iterates data (Row vise), calculate and assign the calculated value to the new columns

        For i = 2 To LastRowTick

       

            If GreatestInc = 0 Then

 

                GreatestInc = ws.Cells(i, 11).Value

                GreatestDec = ws.Cells(i, 11).Value

 

            End If

       

            ' Check if current value is greater than previous value

            If ws.Cells(i, 11).Value >= GreatestInc Then

 

                GreatestInc = ws.Cells(i, 11).Value

                GreatestIncTick = ws.Cells(i, 9).Value

 

            End If

           

            ' Check if current value is less than previous value

            If ws.Cells(i, 11).Value <= GreatestDec Then

 

                GreatestDec = ws.Cells(i, 11).Value

                GreatestDecTick = ws.Cells(i, 9).Value

               

            End If

           

            ' Check if current stock volume is greater then previous value

            If ws.Cells(i, 12).Value > GreatestTotalStockVol Then

 

                GreatestTotalStockVol = ws.Cells(i, 12).Value

                GreatestTotalStockVolTick = ws.Cells(i, 9).Value

              

            End If

 

        Next i

       

        ' Assign the column header

        ws.Range("O2").Value = "Greatest % increase"

        ws.Range("O3").Value = "Greatest % decrease"

        ws.Range("O4").Value = "Greatest total volume"

        ws.Range("P1").Value = "Ticker"

        ws.Range("Q1").Value = "Value"

       

        ' Assign the column values

        ws.Range("P2").Value = GreatestIncTick

        ws.Range("P3").Value = GreatestDecTick

        ws.Range("P4").Value = GreatestTotalStockVolTick

       

        With ws.Range("Q2")

             .Value = GreatestInc

             .Style = "Percent"

             .NumberFormat = "0.00%"

        End With

       

        With ws.Range("Q3")

             .Value = GreatestDec

             .Style = "Percent"

             .NumberFormat = "0.00%"

        End With

       

        'ws.Range("Q2").Value = GreatestInc

        'ws.Range("Q3").Value = GreatestDec

        ws.Range("Q4").Value = GreatestTotalStockVol

    Next ws

 

End Sub