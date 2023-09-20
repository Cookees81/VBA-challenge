# VBA Cod FormatYearly2

Sub FormatYearly2()

    For Each ws In Worksheets
                
        LastYearRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
            For i = 2 To LastYearRow
            
                If ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
                    
                End If
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.Color = vbGreen
                
                End If
                
            Next i
            
    Next ws

End Sub
