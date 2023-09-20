# VBA Code MinMax

Sub MinMax()
    
    For Each ws In Worksheets
    
    Dim xmax As Double
    Dim xmin As Double
    
    Dim tickmin As String
    Dim tickmax As String
    
    Dim r As Range

    Set r = ws.Range("K2:K" & Rows.Count)
    xmin = Application.WorksheetFunction.Min(r)
    xmax = Application.WorksheetFunction.Max(r)
    
        ws.Cells(3, 17).Value = xmin
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(2, 17).Value = xmax
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
         For i = 2 To 125
    
            If ws.Cells(3, 17).Value = ws.Cells(i, 11) Then
                tickmin = ws.Cells(i, 9).Value
                ws.Cells(3, 16).Value = tickmin
            End If
            
            If ws.Cells(2, 17).Value = ws.Cells(i, 11) Then
                tickmax = ws.Cells(i, 9).Value
                ws.Cells(2, 16).Value = tickmax
            End If
        
        Next i
        
    Next ws
        
End Sub