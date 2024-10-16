Sub stock_output_tables()

' -----------------------------------------------
' DESCRIPTION: Generates 2 ouptut tables for each
' quarter to analyze generated stock data
' -----------------------------------------------

' loop through all sheets
For Each ws In Worksheets
    
    ' skip Instructions sheet
    If ws.Name <> "Instructions" Then
    
        ' -------------
        ' set up output
        ' -------------
        
        ' clear previous outputs
        ws.Range("I:L,O:Q").ClearContents
        
        ' create value output column names
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Quarterly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        ' create summary output naming
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        
        ' ---------------------------
        ' populate first output table
        ' ---------------------------
        
        ' get last row of sheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' initialize values
        OutputRow = 2  ' output row
        TickerStart = 2  ' track first ticker row
        StockVolume = 0  ' track stock volume count
        
        ' loop through rows
        For r = 2 To LastRow
        
            ' check if when switches to new stock ticker
            If ws.Cells(r, "A").Value <> ws.Cells(r + 1, "A").Value Then
                ' add Ticker output
                ws.Cells(OutputRow, "I").Value = ws.Cells(r, "A").Value
                
                ' add Quarterly Change output
                OpenPrice = ws.Cells(TickerStart, "C").Value ' opening price
                ClosePrice = ws.Cells(r, "F").Value ' closing price
                ws.Cells(OutputRow, "J").Value = FormatCurrency(ClosePrice - OpenPrice)
                
                ' add Percent Change output
                ws.Cells(OutputRow, "K").Value = FormatPercent((ClosePrice - OpenPrice) / OpenPrice)
                
                ' add Total Stock Volume output
                ws.Cells(OutputRow, "L").Value = ws.Cells(r, "G").Value + StockVolume
                
                ' -------------------
                ' update below values
                OutputRow = OutputRow + 1  ' increase output row
                TickerStart = r + 1  ' update start of next ticker
                StockVolume = 0  ' restart stock volume
          
            ' same stock ticker
            Else
                ' add total stock volume
                StockVolume = ws.Cells(r, "G").Value + StockVolume
                
            End If
            
        Next r
        
        
        ' ----------------------------
        ' populate second output table
        ' ----------------------------
        
        ' get last row in output column
        LastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        
        ' get summary table outputs
        
        ' greatest % increase
        MaxPercentChange = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow))  ' get max change value
        MaxPercentRow = Application.WorksheetFunction.Match(MaxPercentChange, ws.Range("K2:K" & LastRow), 0) + 1  ' get max change row
        ws.Cells(2, "P").Value = ws.Cells(MaxPercentRow, "I").Value ' show associated ticker
        ws.Cells(2, "Q").Value = FormatPercent(MaxPercentChange) ' show max % change
    
        ' greatest % decrease
        MinPercentChange = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow))  ' get min change value
        MinPercentRow = Application.WorksheetFunction.Match(MinPercentChange, ws.Range("K2:K" & LastRow), 0) + 1  ' get min change row
        ws.Cells(3, "P").Value = ws.Cells(MinPercentRow, "I").Value ' show associated ticker
        ws.Cells(3, "Q").Value = FormatPercent(MinPercentChange) ' show min % change
        
        ' greatest total volume
        MaxStock = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow))  ' get max stock total
        MaxStockRow = Application.WorksheetFunction.Match(MaxStock, ws.Range("L2:L" & LastRow), 0) + 1  ' get max stock total row
        ws.Cells(4, "P").Value = ws.Cells(MaxStockRow, "I").Value ' show associated ticker
        ws.Cells(4, "Q").Value = MaxStock ' show max stock volume
        
        
        ' --------------------
        ' format output tables
        ' --------------------
        
        ' conditional formatting for Quarterly Change + Percent Change
        
        ' set up range for conditional formatting
        Set Rng = ws.Range("J2:K" & LastRow)  ' define range
        Rng.FormatConditions.Delete  ' clear any existing conditional formatting
        
        ' apply conditional formatting
        With Rng
            ' positive value then make cell green
            .FormatConditions.Add xlCellValue, xlGreater, "0"
            .FormatConditions(1).Interior.ColorIndex = 4
            
            ' negative value then make cell red and font black
            .FormatConditions.Add xlCellValue, xlLess, "0"
            .FormatConditions(2).Interior.ColorIndex = 3
            .FormatConditions(2).Font.ColorIndex = 1
            
            ' 0 value then make cell no fill
            .FormatConditions.Add xlCellValue, xlEqual, "0"
            .FormatConditions(3).Interior.ColorIndex = xlNone
            
        End With
        
        
        ' adjust columns' width
        ws.Columns("I:Q").AutoFit
        
        ' adjust output titles
        ws.Range("I1:L1,O2:O4,P1:Q1").Font.Bold = True  ' bolden titles
        ws.Range("I1:L1,O2:O4,P1:Q1").Interior.Color = RGB(217, 217, 217)  ' fill with gray
        
        ' add borders
        ws.Range("I1:L" & LastRow).Borders.LineStyle = xlContinuous  ' first output
        ws.Range("O1:Q4").Borders.LineStyle = xlContinuous ' second output
    End If

Next ws


End Sub

