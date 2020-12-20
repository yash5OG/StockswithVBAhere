Sub stockshomework():

'define variables for the code
'ticker symbol, greatest price increase ticker, greatest price decrease ticker, greatest stock volume ticker
Dim ticker, gpi_ticker, gpd_ticker, gsv_ticker  As String

Dim n As Integer

Dim lastRow, lastColumn As Long

'opening price, closing price, change, change %, total sales volume, greatest price increase, greatest price decrease, greatest stock volume
Dim op, cp, delta, delta_per, tsv, gpi, gpd, gsv As Double

' loop through worksheets (#bonus)
For Each ws In Worksheets

    ws.Activate
    'Summary table#1
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'table skeleton
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Change"
    ws.Range("k1").Value = "% Change"
    ws.Range("l1").Value = "Total Stock Volume"
    
    n = 0
    ticker = ""
    delta = 0
    op = 0
    delta_per = 0
    tsv = 0
    
    'for looping through the data
    For i = 2 To lastRow
        'giving values to variables
        ticker = Cells(i, 1).Value
        If op = 0 Then
            op = Cells(i, 3).Value
        End If
        
        ' cumulating the total value.
        tsv = tsv + Cells(i, 7).Value
        
        ' Run this if we get to a different ticker in the list.
        If Cells(i + 1, 1).Value <> ticker Then
            'new ticker
            n = n + 1
            Cells(n + 1, 9) = ticker
            
            cp = Cells(i, 6)

            'calculations            
            delta = cp - op
            Cells(n + 1, 10).Value = delta
            
            'conditional formatting
            If delta > 0 Then
                Cells(n + 1, 10).Interior.ColorIndex = 4
            
            ElseIf delta < 0 Then
                Cells(n + 1, 10).Interior.ColorIndex = 3
            
            Else
                Cells(n + 1, 10).Interior.ColorIndex = 6
            End If
            
            Cells(n + 1, 12).Value = tsv
            
            ' Calculate percent change value for ticker.
            If op = 0 Then
                delta_per = 0
            Else
                delta_per = (delta / op)
            End If
                        
            Cells(n + 1, 11).Value = Format(delta_per, "Percent")
            
            'to reset
            op = 0 
            tsv = 0
        End If
        
    Next i
    
    'Summary table#2
    Range("o2").Value = "Greatest % Increase"
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "Greatest Total Volume"
    Range("p1").Value = "Ticker"
    Range("q1").Value = "Value"
    
    lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Table structure
    gpi = Cells(2, 11).Value
    g_ticker = Cells(2, 9).Value
    gpd = Cells(2, 11).Value
    gpd_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    'loop to find values
    For i = 2 To lastRow
        If Cells(i, 11).Value > gpi Then
            gpi = Cells(i, 11).Value
            g_ticker = Cells(i, 9).Value
        End If
    
        If Cells(i, 11).Value < gpd Then
            gpd = Cells(i, 11).Value
            gpd_ticker = Cells(i, 9).Value
        End If
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    
    Range("P2").Value = Format(g_ticker, "Percent")
    Range("Q2").Value = Format(gpi, "Percent")
    Range("P3").Value = Format(gpd_ticker, "Percent")
    Range("Q3").Value = Format(gpd, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws


End Sub
