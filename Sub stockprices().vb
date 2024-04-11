Sub stockprices()

    Dim sheet As Worksheet
    Dim i As Long
    Dim OutputRow As Long
    OutputRow = 2
    Dim counter As Integer
    
    
    Dim Ticker_name As String
    Dim YearOpenPrice As Double
    Dim YearClosingPrice As Double
    Dim YearlyChange As Double
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
    
    For Each sheet In Worksheets
    
    
    OutputRow = 2
    
    Dim RowsCount As Long
    RowsCount = sheet.Range("A1").End(xlDown).Row
    
    sheet.Range("I1").EntireColumn.Insert
    sheet.Cells(1, 9).Value = "Ticker"
    
    sheet.Range("J1").EntireColumn.Insert
    sheet.Cells(1, 10).Value = "Yearly Change"
    
    sheet.Range("K1").EntireColumn.Insert
    sheet.Cells(1, 11).Value = "Percent Change"
    
    sheet.Range("L1").EntireColumn.Insert
    sheet.Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To RowsCount
    
        If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
        
            Ticker_name = sheet.Cells(i, 1).Value
            sheet.Range("I" & OutputRow).Value = Ticker_name
            
            YearClosingPrice = sheet.Cells(i, 6).Value
            
            YearlyChange = YearClosingPrice - YearOpenPrice
            sheet.Range("J" & OutputRow).Value = YearlyChange
            
            percentchange = (YearlyChange / YearOpenPrice) * 100
            sheet.Range("K" & OutputRow).Value = percentchange
            
            TotalStockVolume = TotalStockVolume + sheet.Cells(i, 7).Value
            sheet.Range("L" & OutputRow).Value = TotalStockVolume
            
            OutputRow = OutputRow + 1
            
            TotalStockVolume = 0
            
            counter = 0
            
            
        Else
        
            TotalStockVolume = TotalStockVolume + sheet.Cells(i, 7).Value
           
            
            Do Until counter = 1
            YearOpenPrice = sheet.Cells(i, 3).Value
            YearOpenPrice = YearOpenPrice
            counter = 1
            
        
            Loop
            
             
        End If
        
    Next i
    
Next sheet


            

End Sub

Sub percentages()

If percentchange > max_increase Then
            max_increase = percentchange
            max_increase_ticker = ticker
            
        ElseIf percentchange < max_decrease Then
            max_decrease = percentchange
            max_decrease_ticker = ticker
            
        End If
        
        If TotalVolume > max_volume Then
            max_volume = TotalVolume
            max_volume_ticker = ticker
        End If
            output_table_row = output_table_row + 1
            TotalVolume = 0
            TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
            
            
            
        ws.Range("P2").Value = max_increase_ticker
        ws.Range("P3").Value = max_decrease_ticker
        ws.Range("P4").Value = max_volume_ticker
        ws.Range("Q2").Value = max_increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = max_decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = max_volume
   End If
 
Next Row
            
End Sub


