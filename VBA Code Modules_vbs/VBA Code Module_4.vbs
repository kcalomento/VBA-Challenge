Attribute VB_Name = "Module4"
Sub greatest_values()

    Dim last As Long
    Dim r As Long
    Dim increase As String
    Dim decrease As String
    Dim volume As String
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    Dim percent_change As Double
    Dim total_volume As Double
    
    ' how to get greatest ; switch i&d for lowest
    max_increase = -1
    max_decrease = 1
    max_volume = 0
  
    ' last row in the dataset (tickers in I)
    last = WorksheetFunction.CountA(Range("I:I"))
    
    For r = 2 To last
        ' column K (percent change)
        percent_change = Cells(r, 11).Value
        ' column L (total stock volume)
        total_volume = Cells(r, 12).Value
        
        ' greatest % increase
        If percent_change > max_increase Then
            max_increase = percent_change
            ' ticker
            increase = Cells(r, 9).Value
        End If
        
        ' greatest % decrease
        If percent_change < max_decrease Then
            max_decrease = percent_change
            ' ticker
            decrease = Cells(r, 9).Value
        End If
        
        ' Check for greatest total volume
        If total_volume > max_volume Then
            max_volume = total_volume
            ' ticker
            volume = Cells(r, 9).Value
        End If
        
Next r
    
    ' outputs for greatest in specific cells
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    
    ' ticker: greatest increase
    Range("O2").Value = increase
    ' ticker: greatest decrease
    Range("O3").Value = decrease
    ' ticker: greatest total volume
    Range("O4").Value = volume
    
    ' greatest % increase
    Range("P2").Value = max_increase
    ' greatest % decrease
    Range("P3").Value = max_decrease
    ' greatest total volume
    Range("P4").Value = max_volume
    
    
End Sub

