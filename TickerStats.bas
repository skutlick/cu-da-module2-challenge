Attribute VB_Name = "Module1"
Sub TickerStats()
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim QuarterlyPriceChange As Double
    Dim QuarterlyPercentPriceChange As Double
    Dim PriorDateHelper As Long
    PriorDateHelper = 2
    Dim GreatestPriceIncrease As Double
    GreatestPriceIncrease = 0
    Dim GreatestPriceDecrease As Double
    GreatestPriceDecrease = 0
    Dim TotalVolume As Double
    TotalVolume = 0
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = 0
    Dim StatsTableRow As Long
    StatsTableRow = 2
    Dim LastTickerRow As Long
    LastTickerRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim LastStatsTableRow As Long
    
    For i = 2 To LastTickerRow
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                ws.Range("I" & StatsTableRow).Value = Ticker
                
                ws.Range("L" & StatsTableRow).Value = TotalVolume
                
                TotalVolume = 0
            
                OpenPrice = ws.Range("C" & PriorDateHelper)
                
                ClosePrice = ws.Range("F" & i)
                
                QuarterlyPriceChange = ClosePrice - OpenPrice
                ws.Range("J" & StatsTableRow).Value = QuarterlyPriceChange
                
                If OpenPrice = 0 Then
                    PercentPriceChange = 0
                    
                    Else
                    QuarterlyPercentPriceChange = QuarterlyPriceChange / OpenPrice
                    
                End If
                
                ws.Range("K" & StatsTableRow).Value = QuarterlyPercentPriceChange
                ws.Range("K" & StatsTableRow).NumberFormat = "0.00%"
                
                If ws.Range("J" & StatsTableRow).Value >= 0 Then
                    ws.Range("J" & StatsTableRow).Interior.ColorIndex = 4
                    
                    Else
                    ws.Range("J" & StatsTableRow).Interior.ColorIndex = 3
                
                End If
                
                StatsTableRow = StatsTableRow + 1
                
                PriorDateHelper = i + 1
                
            End If
    
    Next i
    
    LastStatsTableRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To LastStatsTableRow
    
    If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
        ws.Range("Q2").Value = ws.Range("K" & i).Value
        ws.Range("P2").Value = ws.Range("I" & i).Value
    End If
    
    If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
        ws.Range("Q3").Value = ws.Range("K" & i).Value
        ws.Range("P3").Value = ws.Range("I" & i).Value
    End If
    
    If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
        ws.Range("Q4").Value = ws.Range("L" & i).Value
        ws.Range("P4").Value = ws.Range("I" & i).Value
    End If
    
    Next i
    
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ws.Columns("A:Q").AutoFit
    
    Next ws
    
End Sub
