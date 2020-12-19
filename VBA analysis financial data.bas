Attribute VB_Name = "Module1"
' VBA analysis of green stock performance data

Sub StockPerformance()

    'Variables used
    Dim TickerName As String
    Dim TickerSummaryRow As Integer
    Dim TotalStock As Double
    Dim OpenPrice As Double
    Dim YearChange As Double
    Dim i As Long
    
    'Set starting values of variables
    TickerSummaryRow = 2
    TotalStock = 0
    YearChange = 0
    OpenPrice = Cells(2, 3).Value
    

    ' Create headers for columns for new fields (Ticker, Yearly Change, Percent Change, Total Stock Volume)
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    'Ticker and total stock volume steps
    
    ' Loop through rows in the column  (NEED TO USE LAST ROW - THIS MUST CHANGE)
    'Hard coding last row of column A in testing data for now (will change so can find last row in sheet)
    For i = 2 To 70926

        

    ' Searches for when the value of the next cell is different than that of the current cell
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
            TickerName = Cells(i, 1).Value      'Store ticker strings
            
            TotalStock = TotalStock + Cells(i, 7).Value  ' Add volume value to total
            
            YearChange = (Cells(i, 6).Value - OpenPrice)    ' Store yearly change value
            
            MsgBox (YearChange)
        
            Range("I" & TickerSummaryRow).Value = TickerName  ' Puts string of ticker into first available row in column I
        
            Range("J" & TickerSummaryRow).Value = YearChange  ' puts difference of close price at last row with open price of first row
        
            Range("L" & TickerSummaryRow).Value = TotalStock  ' puts sum of stock vlumn into column L
            
            TickerSummaryRow = TickerSummaryRow + 1 ' Increase row number to organise values
            
            TotalStock = 0 'reset before next loop
            
            OpenPrice = Cells(i + 1, 3).Value  'set open price to value in next row (new ticker)
            
        Else
        
            TotalStock = TotalStock + Cells(i, 7).Value
            
        End If
    
    Next i
    
    
    'Adjust width of columns to match contents of cells for readability
    
    Columns("I:L").AutoFit


End Sub
