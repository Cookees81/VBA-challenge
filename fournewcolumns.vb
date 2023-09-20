# VBA Code fournewcolumns

Sub AlphaTest1()

    For Each ws In Worksheets
    
    Dim Ticker As String
        
    Dim TotalStockVol As Double
    TotalStockVol = 0
    
    Dim OpenStock As Double
    OpenStock = 0
    
    Dim CloseStock As Double
    CloseStock = 0
    
    Dim Yearly As Double
    Yearly = 0
        
    Dim PrcntChg As Double
        
    Dim Summary As Integer
    Summary = 2
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Cells(1, 10).Value = "Yearly"
        
        ws.Cells(1, 11).Value = "Percent Change"
        
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To LastRow
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
            
                Ticker = ws.Cells(i, 1).Value
                
                    ws.Cells(Summary, 9).Value = Ticker
                

                    TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
                
                    ws.Cells(Summary, 12).Value = TotalStockVol
                
                TotalStockVol = 0
                
            Else
                TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
        
            End If
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                OpenStock = ws.Cells(i, 3).Value
                
            End If
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                CloseStock = ws.Cells(i, 6).Value
                
                    PrcntChg = ((CloseStock - OpenStock) / OpenStock)
                
                    ws.Cells(Summary, 11).Value = PrcntChg
                    
                    ws.Cells(Summary, 11).NumberFormat = "0.00%"
                    
                    PrcntChg = 0
                    
                    Yearly = (CloseStock - OpenStock)
                
                    ws.Cells(Summary, 10).Value = Yearly
                    
                    Yearly = 0
                    
                    Summary = Summary + 1
                                        
            End If
           
        Next i
    
    Next ws

End Sub
