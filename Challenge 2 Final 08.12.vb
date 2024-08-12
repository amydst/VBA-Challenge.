Sub QuaterlyStocksReview()
    Dim WS As Worksheet
    
    Dim Counter As Double
    
    'Dim TickerName As String
    
    'Dim TickerVolumeTotal As Double
   ' TickerVolumeTotal = 0
    
    'Dim SummaryTableRow As Integer
   'SummaryTableRow = 2
    
    'Dim LastRow As Long
    'LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Dim LastDayClose As Double
    'LastDayClose = 0
    
    'Dim FirstDayOpen As Double
    'FirstDayOpen = 0

    For Each WS In Worksheets
    
        Dim TickerName As String
    
        Dim TickerVolumeTotal As Double
        TickerVolumeTotal = 0
        
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
        
        Dim LastRow As Double
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
       
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Quarterly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
    
        'Loop through all rows, gets each ticker name and adds the volume total
        For Counter = 2 To LastRow
            If WS.Cells(Counter + 1, 1).Value <> WS.Cells(Counter, 1).Value Then
                'Get each ticker name
                TickerName = WS.Cells(Counter, 1).Value
            
                'Add the Volume total for the ticker
                TickerVolumeTotal = TickerVolumeTotal + WS.Cells(Counter, 7)
            
                ' Print the ticker name in the summary table
            
                WS.Cells(SummaryTableRow, 9).Value = TickerName
            
                ' Print the Volume total in the summary table
                WS.Cells(SummaryTableRow, 12).Value = TickerVolumeTotal
            
                'Add a row to the summary table
                SummaryTableRow = SummaryTableRow + 1
                
                'Reset the volume total
                
                TickerVolumeTotal = 0
            
            Else
                'Add the volume to the total
                TickerVolumeTotal = TickerVolumeTotal + WS.Cells(Counter, 7).Value
            
            End If
            
        Next Counter
        
            SummaryTableRow = 2
            Dim BothDays As Integer
            BothDays = 0
            
            Dim LastDayClose As Double
            LastDayClose = 0
            
            Dim FirstDayOpen As Double
            FirstDayOpen = 0
            
            Dim QuarterlyChange As Double
            QuarterlyChange = 0
            
            Dim PercentChange As Double
            PercentChange = 0
        
        'Get the first and last day to calculate quarterly change
        For Counter = 2 To LastRow
                        
            If WS.Cells(Counter - 1, 1).Value <> WS.Cells(Counter, 1).Value Then
                FirstDayOpen = WS.Cells(Counter, 3).Value
                'MsgBox ("FD" + Str(FirstDayOpen))
                WS.Cells(SummaryTableRow, 13).Value = FirsDayOpen
                BothDays = BothDays + 1
            End If
                    
            If WS.Cells(Counter + 1, 1).Value <> WS.Cells(Counter, 1).Value Then
                LastDayClose = WS.Cells(Counter, 6)
                '  MsgBox ("LD" + Str(LastDayClose))
                BothDays = BothDays + 1
            End If
            
            'BothDays checks if I have the value of both the first and the last day I'm looking for
            
            If BothDays = 2 Then
            QuarterlyChange = LastDayClose - FirstDayOpen
            WS.Cells(SummaryTableRow, 10).Value = QuarterlyChange
            'MsgBox ("QC" + Str(QuarterlyChange))
            
            'Calculate percent change and format it to %
            
            PercentChange = (QuarterlyChange / FirstDayOpen)
            WS.Cells(SummaryTableRow, 11).Value = PercentChange
            WS.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
            
            'Fills the cells green if there's a positive change, red if there's a negative change
            
            If PercentChange > 0 Then
                WS.Cells(SummaryTableRow, 11).Interior.ColorIndex = 4
            ElseIf PercentChange < 0 Then
               WS.Cells(SummaryTableRow, 11).Interior.ColorIndex = 3
            End If
            
            
            'Adds a row to the table
            SummaryTableRow = SummaryTableRow + 1
            
            'resets the counter for BothDays
            BothDays = 0
            End If
            
        Next Counter
        
        'Names greatest values rows
        
        WS.Cells(1, 18).Value = "Ticker"
        WS.Cells(1, 19).Value = "Value"
        WS.Cells(2, 17).Value = "Greatest % Increase"
        WS.Cells(3, 17).Value = "Greates % Decrease"
        WS.Cells(4, 17).Value = "Greatest Total Volume"
       
        
        'Creating all my variables to hold the greatest values and the corresponding ticker name
        Dim GOATVolumeValue As Double
        GOATVolumeValue = 0
        
        Dim GOATVolumeName As String
        Dim GOATIncreaseName As String
        Dim GOATDecreaseName As String
        
        Dim GOATIncreaseValue As Double
        GOATIncreaseValue = 0
       
        
        Dim GOATDecreaseValue As Double
        GOATDecreaseValue = 0
        
        Dim STLastRow As Double
        STLastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
        For Counter = 2 To STLastRow
        
            If GOATVolumeValue < WS.Cells(Counter + 1, 12).Value Then
                GOATVolumeValue = WS.Cells(Counter + 1, 12).Value
                GOATVolumeName = WS.Cells(Counter + 1, 9).Value
            End If
            
            If GOATIncreaseValue < WS.Cells(Counter + 1, 11).Value Then
                GOATIncreaseValue = WS.Cells(Counter + 1, 11).Value
                GOATIncreaseName = WS.Cells(Counter + 1, 9).Value
            End If
            If GOATDecreaseValue > WS.Cells(Counter + 1, 11).Value Then
                GOATDecreaseValue = WS.Cells(Counter + 1, 11).Value
                GOATDecreaseName = WS.Cells(Counter + 1, 9).Value
            End If
            
        Next Counter
        
        WS.Cells(4, 18).Value = GOATVolumeName
        WS.Cells(4, 19).Value = GOATVolumeValue
        WS.Cells(3, 18).Value = GOATDecreaseName
        WS.Cells(3, 19).Value = GOATDecreaseValue
        WS.Cells(3, 19).NumberFormat = "0.00%"
        WS.Cells(2, 18).Value = GOATIncreaseName
        WS.Cells(2, 19).Value = GOATIncreaseValue
        WS.Cells(2, 19).NumberFormat = "0.00%"
        
    Next WS
    
End Sub