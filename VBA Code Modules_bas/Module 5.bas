Attribute VB_Name = "Module5"
' THIS IS MODULES 1-4 ON THE SUBMITTED XLSM SHEET COMBINED
' THIS MODULE (5) IS ALL THE VBA CODE FROM 1-4

Sub quarterly_ticker()

    Dim out_row As Long
    Dim r As Long
    Dim last As Long
    
    ' last row in the dataset (tickers in A)
    last = WorksheetFunction.CountA(Range("A:A"))
    
    ' begin output at row 2
    out_row = 2
    
    For r = 2 To last
       ' end of a specific ticker (change to the next)
        If Cells(r, 1).Value <> Cells(r + 1, 1).Value Or r = last Then
            
            Cells(out_row, 9).Value = Cells(r, 1).Value
            
            out_row = out_row + 1
            
        End If
        
    Next r
    
End Sub

Sub quarterly_and_percent_change()

    Dim out_row As Long
    Dim r As Long
    Dim start As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim change_price As Double
    Dim percent_change As Double
    Dim ticker As String
    Dim last As Long
    
    ' last row in the dataset (tickers in A)
    last = WorksheetFunction.CountA(Range("A:A"))
    
    ' begin output at row 2
    out_row = 2
    start = 2
    
    For r = 2 To last
    
        ' end of a specific ticker (change to the next)
        If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
            ' column C is open
            open_price = Cells(start, 3).Value
            
            ' column F is close
            close_price = Cells(r, 6).Value
            
            ' change calculation (Close - Open)
            change_price = close_price - open_price
            
            ' output the change in column J
            Cells(out_row, 10).Value = change_price
            
            ' calculate  percent change
            If open_price <> 0 Then
                percent_change = (change_price / open_price) * 100
            Else
                ' avoid divisble by 0
                percent_change = 0
            End If
            
            ' output percent change in K
            Cells(out_row, 11).Value = percent_change
            
            ' conditional format for change (quaterly change): green positive, red negative
            If change_price > 0 Then
                Cells(out_row, 10).Interior.ColorIndex = 4 ' Green
                
            ElseIf change_price < 0 Then
                Cells(out_row, 10).Interior.ColorIndex = 3 ' Red
                
            End If
                
            ' conditional format for change (percent change): green positive, red negative
            If percent_change > 0 Then
                Cells(out_row, 11).Interior.ColorIndex = 4 ' Green
                
            ElseIf percent_change < 0 Then
                Cells(out_row, 11).Interior.ColorIndex = 3 ' Red
                
            End If
            
            ' next ticker
            out_row = out_row + 1
            
            ' begin again for next
            start = r + 1
            
            ' If last row, exit the loop
            If r = last Then
            Exit For
            
        End If
        
    End If
    
Next r
    
End Sub

Sub total_stock_volume()

    Dim out_row As Long
    Dim r As Long
    Dim start As Long
    Dim total_volume As Double
    Dim ticker As String
    Dim last As Long
    
    ' last row in the dataset (tickers in A)
    last = WorksheetFunction.CountA(Range("A:A"))
    
    ' begin output at row 2
    out_row = 2
    
    start = 2
    
    For r = 2 To last
        
        ' end of a specific ticker (change to the next)
        If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
        
            total_volume = 0
            
            For i = start To r
            
            ' this is for column G
            total_volume = total_volume + Cells(i, 7).Value
            
            Next i
        
            ' output total stock volume in column L
            Cells(out_row, 12).Value = total_volume
            
            ' next ticker
            out_row = out_row + 1
            
            ' begin again for next
            start = r + 1
            
        End If
        
          ' If last row, exit the loop
            If r = last Then
            Exit For
            
        End If
        
    End If
    
Next r

    
End Sub

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

