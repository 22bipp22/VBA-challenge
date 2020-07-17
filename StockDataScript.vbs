Attribute VB_Name = "Module1"
 Sub WorksheetLoop()
 
    Dim sh As Worksheet
    
    For Each sh In Worksheets
        sh.Select
        Call stockChange
    
    Next

    
End Sub


Sub stockChange()
Dim Ticker As String
    Dim FirstOpen As Double
    Dim LastClose As Double
    Dim YearlyChange As Double
    Dim TotalStockVol As Double
    Dim PercentChange As Double
    Dim SummaryTableRow As Integer
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GITicker As String
    Dim GDTicker As String
    Dim GVTicker As String
       
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set up header row for summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Set initial value of variables
    SummaryTableRow = 2
    FirstOpen = Cells(2, 3).Value
    TotalStockVol = Cells(2, 7).Value
    
    For I = 2 To lastrow
        
        If Cells(I, 1).Value = Cells(I + 1, 1).Value Then
            
            TotalStockVol = TotalStockVol + Cells(I, 7).Value
            
            
        Else
            
            Ticker = Cells(I, 1).Value
            TotalStockVol = TotalStockVol + Cells(I, 7).Value
            LastClose = Cells(I, 6).Value
            YearlyChange = LastClose - FirstOpen
                        
            If FirstOpen = 0 Then
                PercentChange = 0 'to avoid divide by zero error on sheet p
            Else
                PercentChange = YearlyChange / FirstOpen
            End If
            
            Cells(SummaryTableRow, 9).Value = Ticker
            Cells(SummaryTableRow, 10).Value = YearlyChange
            Cells(SummaryTableRow, 11).Value = PercentChange
            Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
            Cells(SummaryTableRow, 12).Value = TotalStockVol
           
            'Check for a positive yearly change and color the cell green
            If YearlyChange > 0 Then
            
                Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                
            Else
            
                Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                
            End If
            
            'Check if % Increase is greater than previous greatest
            If PercentChange > GreatestIncrease Then
                
                GITicker = Ticker
                GreatestIncrease = PercentChange
                
            End If
            
            'Check if % Decrease is less than previous least
            If PercentChange < GreatestDecrease Then
            
                GDTicker = Ticker
                GreatestDecrease = PercentChange
                
            End If
            
            'Check for greatest volume
            If TotalStockVol > GreatestVolume Then
                
                GVTicker = Ticker
                GreatestVolume = TotalStockVol
                
            End If
            
                     
            'Set up variables for the next ticker symbol
            SummaryTableRow = SummaryTableRow + 1
            FirstOpen = Cells(I + 1, 3).Value
            TotalStockVol = 0
           
        End If
        
    Next I
           
    'Populate summary with Greatest
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("P2").Value = GITicker
    Range("Q2").Value = GreatestIncrease
    Range("P3").Value = GDTicker
    Range("Q3").Value = GreatestDecrease
    Range("P4").Value = GVTicker
    Range("Q4").Value = GreatestVolume
    
    'Some final formatting
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Columns("I:Q").AutoFit
        
End Sub

