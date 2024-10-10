Attribute VB_Name = "Module3"
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
        
    Next r
    
End Sub

