Sub Moderate()
    
    Dim NewRowCnt As Integer
    Dim TotalVol As Double
    Dim StartYr As Double
    Dim EndYr As Double
    Dim YrChg As Double
    Dim PctChg As Double
        
    
    For Each ws In Worksheets
        
        NewRowCnt = 2
        StartYr = Range("C2").Value
        TotalVol = Range("G2").Value
        ws.Range("I1").Value = "Ticker"
        ws.Range("I1").Font.FontStyle = "Bold"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("J1").Font.FontStyle = "Bold"
        ws.Range("K1").Value = "Pct Change"
        ws.Range("K1").Font.FontStyle = "Bold"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("L1").Font.FontStyle = "Bold"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        For i = 2 To LastRow
        
            If (ws.Range("A" & (i + 1)).Value <> ws.Range("A" & i).Value) Then
                ws.Range("I" & NewRowCnt).Value = ws.Cells(i, 1).Value
                EndYr = ws.Range("F" & i).Value
                YrChg = EndYr - StartYr
                ws.Range("J" & NewRowCnt).Value = YrChg
                ws.Range("J" & NewRowCnt).NumberFormat = "$#,##0.00"
                    If (ws.Range("J" & NewRowCnt).Value > 0) Then
                        ws.Range("J" & NewRowCnt).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & NewRowCnt).Interior.ColorIndex = 3
                    End If
                If (StartYr > 0 And EndYr > 0) Then
                    PctChg = (EndYr - StartYr) / StartYr
                Else
                    PctChg = 0
                End If
                ws.Range("K" & NewRowCnt).Value = PctChg
                ws.Range("K" & NewRowCnt).NumberFormat = "0.00%"
                ws.Range("L" & NewRowCnt).Value = TotalVol
                ws.Range("L" & NewRowCnt).NumberFormat = "#,##0"
                
                TotalVol = ws.Range("G" & i + 1).Value
                StartYr = ws.Range("C" & i + 1).Value
                NewRowCnt = NewRowCnt + 1
            Else
                TotalVol = TotalVol + Range("G" & i).Value
            End If
        
        Next i

    Next ws

End Sub