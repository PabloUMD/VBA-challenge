Sub AnalyzeStocks()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim Volume As Double
    Dim LastRow As Long
    Dim SummaryRow As Integer
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    ' Loop through each worksheet
    For Each ws In Worksheets
        ' Initialize variables
        SummaryRow = 2
        OpenPrice = 0
        ClosePrice = 0
        Volume = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        
        ' Find the last row with data
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Create headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"
        ws.Cells(1, 11).Value = "Quarterly Change ($)"
        ws.Cells(1, 12).Value = "Percent Change (%)"
        
        ' Loop through all rows
        For i = 2 To LastRow
            ' Check if the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
                Volume = Volume + ws.Cells(i, 7).Value
                
                ' Calculate the quarterly change and percent change
                QuarterlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (QuarterlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If
                
                ' Add the data to the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = Volume
                ws.Cells(SummaryRow, 11).Value = QuarterlyChange
                ws.Cells(SummaryRow, 12).Value = PercentChange
                
                ' Apply conditional formatting
                If QuarterlyChange >= 0 Then
                    ws.Cells(SummaryRow, 11).Interior.Color = vbGreen
                Else
                    ws.Cells(SummaryRow, 11).Interior.Color = vbRed
                End If
                
                If PercentChange >= 0 Then
                    ws.Cells(SummaryRow, 12).Interior.Color = vbGreen
                Else
                    ws.Cells(SummaryRow, 12).Interior.Color = vbRed
                End If
                
                ' Check for greatest increase, decrease, and volume
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestIncreaseTicker = Ticker
                End If
                
                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestDecreaseTicker = Ticker
                End If
                
                If Volume > GreatestVolume Then
                    GreatestVolume = Volume
                    GreatestVolumeTicker = Ticker
                End If
                
                ' Move to the next row in the summary table
                SummaryRow = SummaryRow + 1
                
                ' Reset volume for the next ticker
                Volume = 0
                OpenPrice = ws.Cells(i + 1, 3).Value
            Else
                ' Accumulate volume
                Volume = Volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Output the greatest increase, decrease, and volume
        ws.Cells(SummaryRow + 2, 9).Value = "Greatest % Increase"
        ws.Cells(SummaryRow + 2, 10).Value = GreatestIncreaseTicker
        ws.Cells(SummaryRow + 2, 12).Value = GreatestIncrease
        
        ws.Cells(SummaryRow + 3, 9).Value = "Greatest % Decrease"
        ws.Cells(SummaryRow + 3, 10).Value = GreatestDecreaseTicker
        ws.Cells(SummaryRow + 3, 12).Value = GreatestDecrease
        
        ws.Cells(SummaryRow + 4, 9).Value = "Greatest Total Volume"
        ws.Cells(SummaryRow + 4, 10).Value = GreatestVolumeTicker
        ws.Cells(SummaryRow + 4, 12).Value = GreatestVolume
    Next ws
End Sub
