Sub StockLoop()

Dim ws As Worksheet

For Each ws In Worksheets
    
    ws.Activate

    Dim Total As Double
    Dim Counter As Double
    Dim OpenAmount As Double
    Dim CloseAmount As Double
    
    Counter = 2
    OpenAmount = 0
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To LastRow
    
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            
            Total = Total + Cells(i, 7).Value
            
            If OpenAmount = 0 Then
                OpenAmount = Cells(i, 3).Value
            End If
            
            CloseAmount = Cells(i, 6).Value
            
        Else
        
            Total = Total + Cells(i, 7).Value
            CloseAmount = Cells(i, 6).Value
            Cells(Counter, 9).Value = Cells(i, 1).Value
            Cells(Counter, 10).Value = CloseAmount - OpenAmount
            
                If Cells(Counter, 10).Value < 0 Then
                    Cells(Counter, 10).Interior.ColorIndex = 3
                Else
                    Cells(Counter, 10).Interior.ColorIndex = 4
                End If
            
            Cells(Counter, 11).Value = (CloseAmount - OpenAmount) / OpenAmount
            Cells(Counter, 12).Value = Total
            Total = 0
            Counter = Counter + 1
            OpenAmount = 0
            CloseAmount = 0
        End If
    Next i
    
Next

MsgBox ("Loop Complete!")

End Sub

