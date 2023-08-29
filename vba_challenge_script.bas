Attribute VB_Name = "Module1"
Sub StockData()

    For Each ws In Worksheets
    
    Dim Worksheet As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStock As Double
    
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim TIncrease As String
    Dim TDecrease As String
    Dim TVolume As String
    
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
    
    TotalStock = 0
    
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    'to autofit the columns
    ws.Columns(10).AutoFit
    ws.Columns(11).AutoFit
    ws.Columns(12).AutoFit
    ws.Columns(15).AutoFit
    
    
    Dim row As Integer
    row = 2
    
    YearlyChange = 0
    ClosePrice = 0
    OpenPrice = 0
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        
           'ticker
            ws.Range("I" & row) = ws.Cells(i, 1)
            
            'yearly change
            
              OpenPrice = OpenPrice + ws.Cells(i, 3)
              ClosePrice = ClosePrice + ws.Cells(i, 6)
              
              YearlyChange = ClosePrice - OpenPrice
              ws.Range("J" & row) = YearlyChange
                
                
            'conditional formating to fill in the colour for yearly change
                If YearlyChange > 0 Then
                    ws.Range("J" & row).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    ws.Range("J" & row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & row).Interior.ColorIndex = 48
                End If
    
            'percent change
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If
                ws.Range("K" & row) = Format(PercentChange, "0.00%")
            
            'total stock
              TotalStock = TotalStock + ws.Cells(i, 7)
              ws.Range("L" & row).NumberFormat = "0"
              ws.Range("L" & row) = TotalStock
              
              
              'greatest % increase and % decrease
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    TIncrease = ws.Cells(i, 1)
                ElseIf PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    TDecrease = ws.Cells(i, 1)
                End If
                
                ws.Range("P2") = TIncrease
                ws.Range("Q2") = Format(GreatestIncrease, "0.00%")
                
                ws.Range("P3") = TDecrease
                ws.Range("Q3") = Format(GreatestDecrease, "0.00%")
                
                'greatest volume
                 If TotalStock > GreatestVolume Then
                    GreatestVolume = TotalStock
                    TVolume = ws.Cells(i, 1)
                End If
                
                ws.Range("Q4").NumberFormat = "0"
                ws.Range("P4") = TVolume
                ws.Range("Q4") = GreatestVolume
              
              row = row + 1
              
              OpenPrice = 0
              ClosePrice = 0
              TotalStock = 0
              
              Else
              TotalStock = TotalStock + ws.Cells(i, 7)
              OpenPrice = OpenPrice + ws.Cells(i, 3)
              ClosePrice = ClosePrice + ws.Cells(i, 6)
            End If
        
       Next i
             
    Next ws

End Sub



