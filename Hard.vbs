Sub Hard()

Dim RowInc As Integer
Dim RowDec As Integer
Dim RowVol As Integer

    For Each ws In Worksheets
    
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        RowInc = 2
        RowDec = 2
        RowVol = 2
     
        For i = 2 To LastRow
      
            If ws.Range("K" & i).Value > ws.Range("K" & RowInc).Value Then
                RowInc = i
            End If
            
            If ws.Range("K" & i).Value < ws.Range("K" & RowDec).Value Then
                RowDec = i
            End If
            
            If ws.Range("L" & i).Value > ws.Range("L" & RowVol).Value Then
                RowVol = i
            End If
      
        Next i
        
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O2").Font.FontStyle = "Bold"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O3").Font.FontStyle = "Bold"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("O4").Font.FontStyle = "Bold"
        ws.Range("P1").Value = "Ticker"
        ws.Range("P1").Font.FontStyle = "Bold"
        ws.Range("P2").Value = ws.Range("I" & RowInc).Value
        ws.Range("P3").Value = ws.Range("I" & RowDec).Value
        ws.Range("P4").Value = ws.Range("I" & RowVol).Value
        ws.Range("Q1").Value = "Value"
        ws.Range("Q1").Font.FontStyle = "Bold"
        ws.Range("Q2").Value = ws.Range("K" & RowInc).Value
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = ws.Range("K" & RowDec).Value
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = ws.Range("L" & RowVol).Value
        ws.Range("Q4").NumberFormat = "#,##0"
     
    
    Next ws

End Sub
