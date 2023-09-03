Attribute VB_Name = "Module1"
Sub StockData()

For Each ws In Worksheets

    Dim TCounter As Long
    Dim PercentChange As Double
    Dim TotalStock As Double

    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    Dim i As Long
    Dim row As Long
    
    LastRowTicker = ws.Cells(Rows.Count, 1).End(xlUp).row
    
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
    ws.Columns("A:Z").AutoFit
    
    TCounter = 2
    row = 2

    For i = 2 To LastRowTicker
    
    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        ' Ticker
        ws.Cells(TCounter, 9) = ws.Cells(i, 1)
        
        ' Yearly change
        ws.Cells(TCounter, 10) = ws.Cells(i, 6) - ws.Cells(row, 3)
        
        ' Conditional formatting to fill in the color for yearly change
        If ws.Cells(TCounter, 10).Value < 0 Then
            ws.Cells(TCounter, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(TCounter, 10).Interior.ColorIndex = 4
        End If

        ' Percent change
        If ws.Cells(row, 3) <> 0 Then
            PercentChange = ((ws.Cells(i, 6) - ws.Cells(row, 3)) / ws.Cells(row, 3))
            ws.Cells(TCounter, 11) = Format(PercentChange, "0.00%")
        Else
            ws.Cells(TCounter, 11) = Format(0, "0.00%")
        End If

        ' Total stock
        TotalStock = Application.WorksheetFunction.Sum(Range(ws.Cells(row, 7), ws.Cells(i, 7)))
        ws.Cells(TCounter, 12).NumberFormat = "0"
        ws.Cells(TCounter, 12) = TotalStock

        TCounter = TCounter + 1
        row = i + 1
    
    End If
    
    Next i
    
    LastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).row
    
    GreatestIncrease = ws.Cells(2, 11)
    GreatestDecrease = ws.Cells(2, 11)
    GreatestVolume = ws.Cells(2, 12)

        For i = 2 To LastRowSummary

        ' Greatest % increase and % decrease, and volume
            If ws.Cells(i, 12) > GreatestVolume Then
                GreatestVolume = ws.Cells(i, 12)
                ws.Cells(4, 16) = ws.Cells(i, 9)
            End If
            
            If ws.Cells(i, 11) > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11)
                ws.Cells(2, 16) = ws.Cells(i, 9)
            End If
            
            If ws.Cells(i, 11) < GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, 11)
                ws.Cells(3, 16) = ws.Cells(i, 9)
            End If
        Next i
        
        ' Format the result cells
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0"


    Next ws
    
End Sub
