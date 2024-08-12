Sub QuaterlyStocksReview()
    Dim WS As Worksheet
    
    Dim Counter As Integer
    
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
        
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Dim LastDayClose As Double
        'LastDayClose = 0
        
        'Dim FirstDayOpen As Double
        'FirstDayOpen = 0
    
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
            
            If BothDays = 2 Then
            QuarterlyChange = LastDayClose - FirstDayOpen
            WS.Cells(SummaryTableRow, 10).Value = QuarterlyChange
            'MsgBox ("QC" + Str(QuarterlyChange))
            
            PercentChange = (QuarterlyChange / FirstDayOpen)
            WS.Cells(SummaryTableRow, 11).Value = PercentChange
            WS.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
            
            SummaryTableRow = SummaryTableRow + 1
            BothDays = 0
            End If
            
        Next Counter
        
        
        
    Next WS
    
End Sub