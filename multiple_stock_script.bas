Attribute VB_Name = "Module1"
Sub StockAnalysis()

'Loop through all worksheets
For Each ws In Worksheets
    
    'Set variable for holding stock ticker
    Dim stockTicker As String
    
    'Set initial variable for year
    Dim opening As Double
    Dim closing As Double
    
    'Set initial variable for percentage change
    Dim percentChange As Double
    
    'Set initial variable for Stock Volume
    Dim stockVolume As Double
    stockVolume = 0
    
    'Keep track of the location for each ticker in the summary table
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
    'Set greatest percentage increase and decrease
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim increaseTick As String
    Dim decreaseTick As String
    Dim greatestVol As Double
    Dim volTick As String
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVol = 0

    'Set summary column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'Loop through all stocks
    For i = 2 To ws.Cells(Rows.Count, 2).End(xlUp).Row
        
        'Check to see if it's in the same stock, if it is not, then
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
            stockTicker = ws.Cells(i, 1).Value
            
            'Print Ticker symbol to Summary table
            ws.Range("I" & SummaryTableRow).Value = stockTicker
            
            'Print opening value to Summary Table
            If ws.Cells(i - k, 3).Value <> 0 Then
                opening = ws.Cells(i - k, 3).Value
                closing = ws.Cells(i, 6).Value
                ws.Range("J" & SummaryTableRow).Value = closing - opening
                     
                'Print percent change to Summary table
                percentChange = (closing - opening) / opening
                ws.Range("K" & SummaryTableRow).Value = FormatPercent(percentChange, 2)
            Else
                percentChange = 0
                
            End If
            
            'Print the Stock Volume to Summary table
            ws.Range("L" & SummaryTableRow).Value = stockVolume
            
            'Add one to the summary table row
            SummaryTableRow = SummaryTableRow + 1
            
            'Reset the Stock Volume Total
            stockVolume = 0
            
            'Reset counter to 0
            k = 0
            
        Else
            'Set counter for number of days open in each ticket
            k = k + 1
            
            'Add to the Total Stock Volume
            stockVolume = stockVolume + ws.Cells(i, 7).Value
             
        End If
    Next i
    
    'Loop through summary table to color change
    For i = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        
    Next i
    
    'Loop through percentages for increase
    For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row
        'If percentage is greater than the highest value before....
        If ws.Cells(i, 11).Value > greatestIncrease Then
            
            'Set value as the current greatest increase
            greatestIncrease = ws.Cells(i, 11).Value
            increaseTick = ws.Cells(i, 9).Value
        
        Else
        
        End If
        
    Next i
    
    'Loop through percentages for decrease
    For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row
        'If percentage is greater than the highest value before....
        If ws.Cells(i, 11).Value < greatestDecrease Then
            
            'Set value as the current greatest increase
            greatestDecrease = ws.Cells(i, 11).Value
            decreaseTick = ws.Cells(i, 9).Value
        
        Else
        
        End If
        
    Next i
    
    'Loop through stock value
    For i = 2 To ws.Cells(Rows.Count, 12).End(xlUp).Row
        'If volume is greater than the highest value before....
        If ws.Cells(i, 12).Value > greatestVol Then
            
            'Set value as the current greatest volume
            greatestVol = ws.Cells(i, 12).Value
            volTick = ws.Cells(i, 9).Value
        
        Else
        
        End If
        
    Next i
    
    'Print summary to sheet
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = increaseTick
    ws.Cells(2, 17).Value = FormatPercent(greatestIncrease)
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = decreaseTick
    ws.Cells(3, 17).Value = FormatPercent(greatestDecrease)
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = volTick
    ws.Cells(4, 17).Value = greatestVol
    
    Next ws

End Sub
      



