Attribute VB_Name = "Module1"
Sub Stocks()

    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim StockVolume As LongLong
    
    'Headers for summary table
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
        
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
    
    
End Sub



