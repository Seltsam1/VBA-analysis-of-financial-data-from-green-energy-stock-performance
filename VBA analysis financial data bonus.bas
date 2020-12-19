Attribute VB_Name = "Module2"
' VBA analysis of green stock performance data  - Bonus
' Version includes addion to run through each worksheet and return 3 additional values

Sub StockPerformanceBonus()

    'Variables used
    Dim TickerName As String
    Dim TickerSummaryRow As Integer
    Dim TotalStock As Double
    Dim OpenPrice As Double
    Dim YearChange As Double
    Dim PercentChange As Double
    Dim LastRow As Long
    Dim LastRow2 As Long
    Dim i As Long
    
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    
    'Set starting values of variables
    TickerSummaryRow = 2
    TotalStock = 0
    YearChange = 0
    PercentChange = 0
    OpenPrice = Cells(2, 3).Value
    

    ' Create headers for columns for new fields (Ticker, Yearly Change, Percent Change, Total Stock Volume)

    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Additional headers for bonus (greatest % increase, % decrease, total volume, ticker, and value)
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    'Find last row in sheet
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row


    'Ticker and total stock volume steps
    
    For i = 2 To LastRow

    ' Searches for when the value of the next cell is different than that of the current cell
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
            TickerName = Cells(i, 1).Value      'Store ticker strings
            
            TotalStock = TotalStock + Cells(i, 7).Value  ' Add volume value to total
            
            YearChange = (Cells(i, 6).Value - OpenPrice)    ' Store yearly change value
            
                If OpenPrice <> 0 Then 'Error on full data where divided by zero, added this in to address problem
                    PercentChange = ((Cells(i, 6).Value - OpenPrice) / OpenPrice)
                Else
                    PercentChange = Cells(i, 6).Value
                End If
        
            Range("I" & TickerSummaryRow).Value = TickerName  ' Puts string of ticker into first available row in column I
        
            Range("J" & TickerSummaryRow).Value = YearChange  ' puts difference of close price at last row with open price of first row into J
            
            
            'conditional to change formatting fill color for yearly change to red if negative, green if positive
                If Range("J" & TickerSummaryRow).Value < 1 Then
                    Range("J" & TickerSummaryRow).Interior.ColorIndex = 3  'red
                Else
                    Range("J" & TickerSummaryRow).Interior.ColorIndex = 4  'green
                End If
                
            
            Range("K" & TickerSummaryRow).Value = PercentChange ' Puts difference of close price and open price divided by open price into K
            
            Range("K" & TickerSummaryRow).NumberFormat = "0.00%"  'Change formatting of value in column K to percentage
        
            Range("L" & TickerSummaryRow).Value = TotalStock  ' puts sum of stock vlumn into column L
            
            TickerSummaryRow = TickerSummaryRow + 1 ' Increase row number to organise values
            
            TotalStock = 0 'reset before next loop
            
            OpenPrice = Cells(i + 1, 3).Value  'set open price to value in next row (new ticker)
            
        Else
        
            TotalStock = TotalStock + Cells(i, 7).Value
            
        End If
    
    Next i
    
    '-------------------------------------
        ' Bonus calculations
    
    
        'Find last row in column K for bonus calculations
    LastRow2 = Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Initialize variables to first values in 2nd row of new summary table
    GreatestIncrease = Cells(2, 11).Value
    GreatestDecrease = Cells(2, 11).Value
    GreatestVolume = Cells(2, 12).Value

    'Iterate through loop comparing stores value with value in next row
    'If greater, then replace stored value
    For i = 2 To LastRow2
    
        If Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = Cells(i, 11).Value
            Cells(2, 17).Value = Cells(i, 11).Value
            Cells(2, 16).Value = Cells(i, 9).Value
        End If
            
        If Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecrease = Cells(i, 11).Value
            Cells(3, 17).Value = Cells(i, 11).Value
            Cells(3, 16).Value = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value > GreatestVolume Then
            GreatestVolume = Cells(i, 12).Value
            Cells(4, 17).Value = Cells(i, 12).Value
            Cells(4, 16).Value = Cells(i, 9).Value
        End If
    
    Next i
    
    'Format for bonus cells to percentage
    
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    
    'Adjust width of columns to match contents of cells for readability
    Columns("I:Q").AutoFit


End Sub

