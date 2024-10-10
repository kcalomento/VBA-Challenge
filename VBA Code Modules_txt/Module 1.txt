Attribute VB_Name = "Module1"
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


