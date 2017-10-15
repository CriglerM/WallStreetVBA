Sub Easy()
    
    Dim NewRowCnt As Integer
    Dim TotalVol As Double
    
    For Each ws In Worksheets
        
        NewRowCnt = 2
        TotalVol = ws.Range("G2").Value
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        For i = 2 To LastRow
        
            If (ws.Range("A" & (i + 1)) <> ws.Range("A" & i)) Then
            'Cells(i + 1, 1) <> Cells(i, 1)) Then
                ws.Range("I" & NewRowCnt).Value = ws.Cells(i, 1).Value
                ws.Range("J" & NewRowCnt).Value = TotalVol
                TotalVol = ws.Range("G" & i + 1).Value
                NewRowCnt = NewRowCnt + 1
            Else
                TotalVol = TotalVol + ws.Range("G" & i).Value
            End If
        
        Next i

    Next ws

End Sub
