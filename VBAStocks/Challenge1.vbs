Attribute VB_Name = "Module2"
Sub Challenge1()

    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim StockVolume As Single

    
    'Headers for summary table
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"

            
    'Keep track of the location for each ticker in the summary
    Dim Summary_Row As Integer
    Summary_Row = 2
    
    'Count the number of rows
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set first open price & stock volume
    OpenPrice = Cells(2, 3).Value
    StockVolume = 0
      
    'Loop through all stocks
    For i = 2 To lastrow
        
        'See if there is a new ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1).Value
            
            StockVolume = StockVolume + Cells(i, 7).Value
            
            'Get ClosePrice
            ClosePrice = Cells(i, 6).Value
            
            'Calculate Yearly Change
            YearlyChange = ClosePrice - OpenPrice
            
            'Calculate Percent Change
            
            If OpenPrice <> 0 Then
                PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)
            Else
                PercentChange = 0
            End If
            
            
            'Put data in Summary
            Range("I" & Summary_Row).Value = Ticker
            Range("J" & Summary_Row).Value = YearlyChange
            Range("K" & Summary_Row).Value = FormatPercent(PercentChange)
            Range("L" & Summary_Row).Value = StockVolume
            
            'Conditional formatting for yearly change
            If YearlyChange < 0 Then
                
                Range("J" & Summary_Row).Interior.ColorIndex = 3
                
            Else
            
                Range("J" & Summary_Row).Interior.ColorIndex = 4
                
            End If
            
                                    
            'Reset variables for new ticker
            OpenPrice = Cells(i + 1, 3).Value
            StockVolume = 0
            Summary_Row = Summary_Row + 1
    
        Else
        
            StockVolume = StockVolume + Cells(i, 7).Value
                
        End If
        
    Next i
    
    Dim GreatestIncreaseTicker As String
    Dim GreatestIncreasePercent As Double
    Dim GreatestDecreaseTicker As String
    Dim GreatestDecreasePercent As Double
    Dim GreatestVolumeTicker As String
    Dim GreatestVolume As Single
    
    'Count rows of summary table
    LastSummaryRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    GreatestIncreasePercent = Cells(2, 11).Value
    GreatestDecreasePercent = Cells(2, 11).Value
    GreatestVolume = Cells(2, 12).Value

        
    'Loop through summary table
    For j = 2 To LastSummaryRow
    
        'Find Greatest Percent Increase & Decrease
        If Cells(j, 11).Value >= GreatestIncreasePercent Then
           GreatestIncreaseTicker = Cells(j, 9).Value
           GreatestIncreasePercent = Cells(j, 11).Value
        
        ElseIf Cells(j, 11).Value < GreatestDecreasePercent Then
        
           GreatestDecreaseTicker = Cells(j, 9).Value
           GreatestDecreasePercent = Cells(j, 11).Value
        
        End If
        
        
        'Find Greatest Volume
        If Cells(j, 12).Value >= GreatestVolume Then
           GreatestVolumeTicker = Cells(j, 9).Value
           GreatestVolume = Cells(j, 12).Value
                 
        End If
        
    
    Next j
    
    'Set value obtained from loop
    Cells(2, 16) = GreatestIncreaseTicker
    Cells(2, 17) = FormatPercent(GreatestIncreasePercent)
    Cells(3, 16) = GreatestDecreaseTicker
    Cells(3, 17) = FormatPercent(GreatestDecreasePercent)
    Cells(4, 16) = GreatestVolumeTicker
    Cells(4, 17) = GreatestVolume
    
    
    End Sub
    
    

