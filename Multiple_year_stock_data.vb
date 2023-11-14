Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim prevTicker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim dataRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double

    For Each ws In ThisWorkbook.Worksheets
        ' Reset variables for each worksheet
        greatestIncrease = -1
        greatestDecrease = 1
        greatestVolume = -1
        dataRow = 2
        openingPrice = 0
        totalVolume = 0
        prevTicker = ""

        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value

            If ticker <> prevTicker And i > 2 Then
                closingPrice = ws.Cells(i - 1, 6).Value
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = (yearlyChange / openingPrice)
                Else
                    percentChange = 0
                End If

                ws.Cells(dataRow, 9).Value = prevTicker
                ws.Cells(dataRow, 10).Value = yearlyChange
                ws.Cells(dataRow, 11).Value = percentChange
                ws.Cells(dataRow, 11).NumberFormat = "0.00%"
                ws.Cells(dataRow, 12).Value = totalVolume

                If yearlyChange > 0 Then
                    ws.Cells(dataRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(dataRow, 10).Interior.Color = RGB(255, 0, 0)
                End If

                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    ws.Cells(2, 16).Value = prevTicker
                    ws.Cells(2, 17).Value = greatestIncrease
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    ws.Cells(3, 16).Value = prevTicker
                    ws.Cells(3, 17).Value = greatestDecrease
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    ws.Cells(4, 16).Value = prevTicker
                    ws.Cells(4, 17).Value = greatestVolume
                End If

                openingPrice = ws.Cells(i, 3).Value
                totalVolume = 0
                dataRow = dataRow + 1
            End If

            totalVolume = totalVolume + ws.Cells(i, 7).Value
            If prevTicker = "" Or prevTicker <> ticker Then
                openingPrice = ws.Cells(i, 3).Value
            End If
            prevTicker = ticker
        Next i

        closingPrice = ws.Cells(lastRow, 6).Value
        yearlyChange = closingPrice - openingPrice
        If openingPrice <> 0 Then
            percentChange = (yearlyChange / openingPrice)
        Else
            percentChange = 0
        End If

        ws.Cells(dataRow, 9).Value = ticker
        ws.Cells(dataRow, 10).Value = yearlyChange
        ws.Cells(dataRow, 11).Value = percentChange
        ws.Cells(dataRow, 11).NumberFormat = "0.00%"
        ws.Cells(dataRow, 12).Value = totalVolume

    Next ws
End Sub
