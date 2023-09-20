# VBA Code AlphaLastTest

Sub AlphaTestlast()
    
    For Each ws In Worksheets

    Dim vmax As Double
    
    Dim GreatVol As String
    
    Dim r As Range
            
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
            Set r = ws.Range("L2:L" & Rows.Count)
            vmax = Application.WorksheetFunction.Max(r)
            
                ws.Cells(4, 17).Value = vmax
                
                For i = 2 To 3001
                
                    If ws.Cells(4, 17).Value = ws.Cells(i, 12) Then
                        GreatVol = ws.Cells(i, 9).Value
                        ws.Cells(4, 16).Value = GreatVol
                    End If
                    
                Next i
        
    Next ws
    
End Sub
