Attribute VB_Name = "Module1"
Sub MultiYearStock()

    Dim i As Long
    Dim totalVolume As Double
    Dim summaryRow As Integer
    Dim summaryArray(3, 2) As Variant
    Dim red As Integer
    Dim green As Integer
    Dim stTickerCol As Integer
    Dim stYearlyChangeCol As Integer
    Dim stPercentChangeCol As Integer
    Dim stTotalStockVolCol As Integer
    Dim gsTickerCol As Integer
    Dim gsValueCol As Integer
    
    
    ' Global variables
    red = 3
    green = 4
    totalVolume = 0
    ' Data Columns
    dTickerCol = 1
    dOpenCol = 3
    dCloseCol = 6
    dVolCol = 7
    ' Summary Total Columns
    stTickerCol = 9
    stYearlyChangeCol = 10
    stPercentChangeCol = 11
    stTotalStockVolCol = 12
    ' Greatest Summary Columns
    gsTickerCol = 16
    gsValueCol = 17
    
    ' Loop through all worksheets
    For Each ws In Worksheets
        summaryRow = 2
        ' Initialize Summary array
        For i = 0 To 2
            summaryArray(i, 0) = ""
            summaryArray(i, 1) = 0
            If (i = 1) Then
                summaryArray(i, 1) = 99999999
            End If
        Next i
        
        ' ADDING NEW COLUMNS NAMES
        ' ticker
        ws.Range("I1").Value = "Ticker"
         ' Yearly Change
        ws.Range("J1").Value = "Yearly Change"
         ' Percent Change
        ws.Range("K1").Value = "Percent Change"
         ' Total Stock Volume
        ws.Range("L1").Value = "Total Stock Volume"
        ' BONUSColumns
        ' Legend of Greatest Summary
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ' Ticker in Greatest Summary
        ws.Range("P1").Value = "Ticker"
        ' Value in Greatest Summary
        ws.Range("Q1").Value = "Value"
        
        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Initializing opening price
        openingPrice = ws.Cells(2, dOpenCol).Value
        
        ' Loop through all rows
        For i = 2 To lastRow
            totalVolume = totalVolume + ws.Cells(i, dVolCol).Value
            If ws.Cells(i, dTickerCol).Value <> ws.Cells(i + 1, dTickerCol).Value Then
                ticker = ws.Cells(i, dTickerCol).Value
                ' Put current ticker name in summary table
                ws.Cells(summaryRow, stTickerCol).Value = ticker
                ' Put totalvolume over in our summary table
                ws.Cells(summaryRow, stTotalStockVolCol).Value = totalVolume
                ' Getting closing price per row
                closingPrice = ws.Cells(i, dCloseCol).Value
                ' Getting yearly change
                yearlyChange = closingPrice - openingPrice
                ' Put yearly change over in our summary table
                ws.Cells(summaryRow, stYearlyChangeCol).Value = yearlyChange
                ' highlight positive change in green and negative change in red
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, stYearlyChangeCol).Interior.ColorIndex = green
                Else
                    ws.Cells(summaryRow, stYearlyChangeCol).Interior.ColorIndex = red
                End If
                ' Calculate percentage change
                If openingPrice > 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                ' Put percent change over in our summary table
                ws.Cells(summaryRow, stPercentChangeCol).Value = percentChange
                ' BONUS
                ' Determining the greatest increase, decrease and total volume
                If (percentChange > summaryArray(0, 1)) Then
                    summaryArray(0, 0) = ticker
                    summaryArray(0, 1) = percentChange
                End If
                If (percentChange < summaryArray(1, 1)) Then
                    summaryArray(1, 0) = ticker
                    summaryArray(1, 1) = percentChange
                End If
                If (totalVolume > summaryArray(2, 1)) Then
                    summaryArray(2, 0) = ticker
                    summaryArray(2, 1) = totalVolume
                End If
                ' Resetting opening price to the first opening price for the next ticker
                openingPrice = ws.Cells(i + 1, dOpenCol).Value

                ' Add one to the summaryRow
                summaryRow = summaryRow + 1
      
                ' Reset the totalVolume
                totalVolume = 0
            End If

        Next i
        ' Formatting the percent change to percent with two decimals
        ws.Range("K1").EntireColumn.NumberFormat = "0.00%"
        ' BONUS
        For i = 0 To 2
            ws.Cells(i + 2, gsTickerCol).Value = summaryArray(i, 0)
            ws.Cells(i + 2, gsValueCol).Value = summaryArray(i, 1)
            If i < 2 Then
                ws.Cells(i + 2, gsValueCol).NumberFormat = "0.00%"
            End If
        Next i
        
        ' Autofit to display data
        ws.Columns("A:Q").AutoFit
    Next ws
    MsgBox ("Completed")

End Sub
