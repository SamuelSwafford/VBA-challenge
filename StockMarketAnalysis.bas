Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()

    ' Declare variables for the loop
    Dim ws As Worksheet
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim startPrice As Double
    Dim endPrice As Double
    Dim row As Long
    Dim lastRow As Long
    Dim summaryTableRow As Integer
    
    ' Variables for tracking the greatest increase, decrease, and total volume
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    
    Application.ScreenUpdating = False ' Turn off screen updating to improve performance
    
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize the starting values
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        totalVolume = 0
        summaryTableRow = 2
        
        ' Add headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize startPrice with the first opening price of the year
        startPrice = ws.Cells(2, 3).Value
        
        ' Find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        
        ' Loop through all rows of data
        For row = 2 To lastRow
            ' Check if this is the last row of the current ticker
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                ticker = ws.Cells(row, 1).Value
                endPrice = ws.Cells(row, 6).Value
                yearlyChange = endPrice - startPrice
                ' Avoid division by zero error
                If startPrice <> 0 Then
                    percentChange = (yearlyChange / startPrice)
                Else
                    percentChange = 0
                End If
                totalVolume = totalVolume + ws.Cells(row, 7).Value
                
                ' Output the data into the summary table
                ws.Cells(summaryTableRow, 9).Value = ticker
                ws.Cells(summaryTableRow, 10).Value = yearlyChange
                ws.Cells(summaryTableRow, 11).Value = percentChange
                ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
                ws.Cells(summaryTableRow, 12).Value = totalVolume
                
                ' Apply conditional formatting to Yearly Change column (Column J)
                With ws.Range(ws.Cells(summaryTableRow, 10), ws.Cells(summaryTableRow, 10))
                    .FormatConditions.Delete
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                    .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 255, 0) ' Green for positive
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                    .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0) ' Red for negative
                End With
                
                ' Apply conditional formatting to Percent Change column (Column K)
                With ws.Range(ws.Cells(summaryTableRow, 11), ws.Cells(summaryTableRow, 11))
                    .FormatConditions.Delete
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                    .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 255, 0) ' Green for positive
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                    .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0) ' Red for negative
                End With
                
                ' Check for greatest increase, decrease, and volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    tickerGreatestIncrease = ticker
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    tickerGreatestDecrease = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    tickerGreatestVolume = ticker
                End If
                
                ' Prepare for next ticker
                totalVolume = 0
                summaryTableRow = summaryTableRow + 1
                ' Update startPrice for the next ticker
                If row + 1 <= lastRow Then
                    startPrice = ws.Cells(row + 1, 3).Value
                End If
            Else
                totalVolume = totalVolume + ws.Cells(row, 7).Value
            End If
        Next row
        
        ' Output the greatest % increase, % decrease, and total volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = tickerGreatestIncrease
        ws.Cells(3, 16).Value = tickerGreatestDecrease
        ws.Cells(4, 16).Value = tickerGreatestVolume
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 17).Value = greatestVolume
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).NumberFormat = "0.00E+00" ' Format for volume in scientific notation
        
    Next ws
    
    Application.ScreenUpdating = True ' Turn on screen updating again

End Sub

