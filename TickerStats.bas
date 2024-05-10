Attribute VB_Name = "Module1"
Sub TickerStats()

    'Declaring "ws" as Worksheet
    Dim ws As Worksheet
    
    'Beginning of loop through each worksheet
    For Each ws In Worksheets
    
    'Creation of column headers for initial quarterly stats table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Creation of row and column labels for additional stats table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Declaring variables to be used throughout the code, as well as establishing initial variable amounts and/or references
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
    
    'Beginning of loop through rows of Ticker data
    For i = 2 To LastTickerRow
        
        'Equation used to sum Ticker volume
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
            'Logic used to determine if previous row's Ticker is the same as the current row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Establishes the first Ticker used in initial quarterly stats table and places values on the table
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & StatsTableRow).Value = Ticker
                ws.Range("L" & StatsTableRow).Value = TotalVolume
                
                'Resets TotalVolume counter so it can be used in next iteration
                TotalVolume = 0
            
                'References Open and Closing prices, calculates $-price change and %-change, places values on the table and adjusts cell formatting
                OpenPrice = ws.Range("C" & PriorDateHelper)
                ClosePrice = ws.Range("F" & i)
                
                QuarterlyPriceChange = ClosePrice - OpenPrice
                ws.Range("J" & StatsTableRow).Value = QuarterlyPriceChange
                ws.Range("J" & StatsTableRow).NumberFormat = "0.00"
                
                If OpenPrice = 0 Then
                    PercentPriceChange = 0
                    
                    Else
                    QuarterlyPercentPriceChange = QuarterlyPriceChange / OpenPrice
                    
                End If
                
                ws.Range("K" & StatsTableRow).Value = QuarterlyPercentPriceChange
                ws.Range("K" & StatsTableRow).NumberFormat = "0.00%"
                
                If ws.Range("J" & StatsTableRow).Value > 0 Then
                    ws.Range("J" & StatsTableRow).Interior.ColorIndex = 4
                    
                    ElseIf ws.Range("J" & StatsTableRow).Value = 0 Then
                        ws.Range("J" & StatsTableRow).Interior.ColorIndex = 2
                    
                    Else
                    ws.Range("J" & StatsTableRow).Interior.ColorIndex = 3
                
                End If
                
                'Adds a row to the StatsTable
                StatsTableRow = StatsTableRow + 1
                
                'Advances the reference / helper variable to move within a Ticker's subset of data
                PriorDateHelper = i + 1
                
            End If
    
    'Advances to the next row of Ticker data
    Next i
    
    'Determines the last row of the initial quarterly stats table
    LastStatsTableRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'Creates loop to cycle through rows of initial quartelry stats table
    For i = 2 To LastStatsTableRow
    
    'Logic to calculate and place summary stats to a dedicated table
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
    
    'Adjusts formatting on summary stats table
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'AutoFits columns to improve readability
    ws.Columns("A:Q").AutoFit
    
    Next ws
    
End Sub
