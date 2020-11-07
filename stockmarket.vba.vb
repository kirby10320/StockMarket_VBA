Sub StockMarket_VBA():

    Dim Total As Double
    Dim Change As Double
    Dim PercentChange As Double
    Dim J As Long
    Dim WS_Count As Integer
    Dim L As Integer
    
For Each ws In Worksheets
    J = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    RCount = 2
    
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
 
    
        For i = 2 To RowCount
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Total = Total + ws.Cells(i, 7).Value
            
                If ws.Cells(J, 3) = 0 Then
            
                    For nonZeroValue = J To i
                    
                        If ws.Cells(nonZeroValue, 3).Value <> 0 Then
                            J = nonZeroValue
                            Exit For
                        
                        End If
                    
                    Next nonZeroValue
                
                End If
            
                Change = ws.Cells(i, 6).Value - ws.Cells(J, 3).Value
                PercentChange = (Change / ws.Cells(J, 3).Value) * 100
                                   
                ws.Range("I" & RCount).Value = ws.Cells(i, 1).Value
                ws.Range("J" & RCount).Value = Change
                ws.Range("K" & RCount).Value = Round(PercentChange, 2)
                ws.Range("L" & RCount).Value = Total
            
                If Change > 0 Then
                    ws.Range("J" & RCount).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & RCount).Interior.ColorIndex = 3
                End If
            
                Total = 0
                RCount = RCount + 1
                Change = 0
            
            
            Else
                Total = Total + ws.Cells(i, 7).Value
            
            End If
            
    
        Next i
    Next ws
    
End Sub
