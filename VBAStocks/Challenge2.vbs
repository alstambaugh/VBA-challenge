Attribute VB_Name = "Module3"
Sub Challenge2()

    For Each ws In Worksheets
    
        Dim Ticker As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim StockVolume As Single

    
        'Headers for summary table
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"

            
        'Keep track of the location for each ticker in the summary
        Dim Summary_Row As Integer
        Summary_Row = 2
    
        'Count the number of rows
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Set first open price & stock volume
        OpenPrice = ws.Cells(2, 3).Value
        StockVolume = 0
      
        'Loop through all stocks
        For i = 2 To lastrow
        
            'See if there is a new ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                Ticker = ws.Cells(i, 1).Value
            
                StockVolume = StockVolume + ws.Cells(i, 7).Value
            
                'Get ClosePrice
                ClosePrice = ws.Cells(i, 6).Value
            
                'Calculate Yearly Change
                YearlyChange = ClosePrice - OpenPrice
            
                'Calculate Percent Change
            
                If OpenPrice <> 0 Then
                    PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)
                Else
                    PercentChange = 0
                End If
            
            
                'Put data in Summary
                ws.Range("I" & Summary_Row).Value = Ticker
                ws.Range("J" & Summary_Row).Value = YearlyChange
                ws.Range("K" & Summary_Row).Value = FormatPercent(PercentChange)
                ws.Range("L" & Summary_Row).Value = StockVolume
            
                'Conditional formatting for yearly change
                If YearlyChange < 0 Then
                
                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                
                Else
            
                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                
                End If
            
                                    
                'Reset variables for new ticker
                OpenPrice = ws.Cells(i + 1, 3).Value
                StockVolume = 0
                Summary_Row = Summary_Row + 1
    
            Else
        
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                
            End If
        
        Next i
    
        Dim GreatestIncreaseTicker As String
        Dim GreatestIncreasePercent As Double
        Dim GreatestDecreaseTicker As String
        Dim GreatestDecreasePercent As Double
        Dim GreatestVolumeTicker As String
        Dim GreatestVolume As Single
    
        'Count rows of summary table
        LastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        GreatestIncreasePercent = ws.Cells(2, 11).Value
        GreatestDecreasePercent = ws.Cells(2, 11).Value
        GreatestVolume = ws.Cells(2, 12).Value

        
        'Loop through summary table
        For j = 2 To LastSummaryRow
    
            'Find Greatest Percent Increase & Decrease
            If ws.Cells(j, 11).Value >= GreatestIncreasePercent Then
                GreatestIncreaseTicker = ws.Cells(j, 9).Value
                GreatestIncreasePercent = ws.Cells(j, 11).Value
        
            ElseIf ws.Cells(j, 11).Value < GreatestDecreasePercent Then
        
                GreatestDecreaseTicker = ws.Cells(j, 9).Value
                GreatestDecreasePercent = ws.Cells(j, 11).Value
        
            End If
        
        
            'Find Greatest Volume
            If ws.Cells(j, 12).Value >= GreatestVolume Then
                GreatestVolumeTicker = ws.Cells(j, 9).Value
                GreatestVolume = ws.Cells(j, 12).Value
                 
            End If
        
    
        Next j
    
        'Set value obtained from loop
        ws.Cells(2, 16) = GreatestIncreaseTicker
        ws.Cells(2, 17) = FormatPercent(GreatestIncreasePercent)
        ws.Cells(3, 16) = GreatestDecreaseTicker
        ws.Cells(3, 17) = FormatPercent(GreatestDecreasePercent)
        ws.Cells(4, 16) = GreatestVolumeTicker
        ws.Cells(4, 17) = GreatestVolume
    
    Next ws
    
    End Sub
    
  
