Attribute VB_Name = "Module2"
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
        If Cells(r, 1).Value <> Cells(r + 1, 1).Value Or r = last Then
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


