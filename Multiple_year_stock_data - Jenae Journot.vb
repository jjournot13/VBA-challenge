Sub StockPerformance():

 ' Loop through all sheets
  For Each ws In Worksheets

    ' Set an initial variable for holding each ticker symbol
    Dim ticker_symbol As String

    ' Set an initial variable for holding the total of each stock
    Dim year_open_price As Double
    year_open_price = 0

    Dim year_close_price As Double
    year_close_price = 0
  
    Dim stock_volume As LongLong
    stock_volume = 0
    
    ' Create variables for summary data
    Dim increase_number As Double
        
    Dim decrease_number As Double
        
    Dim volume_number As LongLong
    
    ' Create a variable for last row
    Dim LastRow As Long
  
    ' Determine the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Add headers for the Ticker, Yearly Change, Percent Change and Total Stock Volume
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' Keep track of the location for each stock in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
          
    ' Keep track of the location for greatest % of increase and decrease, and total volume
    Dim RowCount As Integer
    RowCount = 2
    
    ' Loop through all stock changes
    For i = 2 To LastRow

      ' Set the new ticker symbol
      ticker_symbol = ws.Cells(i, 1).Value

        ' Get new open value only if open value is set to 0
        If year_open_price = 0 Then
        year_open_price = ws.Cells(i, 3).Value
        End If

      ' Get new close value
      year_close_price = ws.Cells(i, 6).Value
    
      ' Calculate stock volume
      stock_volume = stock_volume + ws.Cells(i, 7).Value

      ' Check if still within the same stock, if it is not...
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

          ' Print the ticker symbol in the summary table
          ws.Range("I" & Summary_Table_Row).Value = ticker_symbol
      
          ' Calculate yearly change
          ws.Range("J" & Summary_Table_Row).Value = year_close_price - year_open_price
          
          ' Calculate percent change
          ws.Range("K" & Summary_Table_Row).Value = (year_close_price - year_open_price) / year_open_price

          ' Calculate stock volume
          ws.Range("L" & Summary_Table_Row).Value = stock_volume

          ' Add one to the summary table row to index to the next row in Excel
          Summary_Table_Row = Summary_Table_Row + 1
      
          ' Reset the year opening price to 0
          year_open_price = 0

          ' Reset the year closing price to 0
          year_close_price = 0

          ' Reset the stock volume to 0
          stock_volume = 0

      End If

    Next i

      ' Add conditional formatting to Yearly and Percent Change
      ws.Columns(10).NumberFormat = "0.00"
      
      With ws.Columns(10).FormatConditions _
      .Add(xlCellValue, xlGreater, "=0")
      With .Interior
      .ColorIndex = 4
      End With
      End With
      
      With ws.Columns(10).FormatConditions _
      .Add(xlCellValue, xlLess, "=0")
      With .Interior
      .ColorIndex = 3
      End With
      End With

      ws.Columns(11).NumberFormat = "0.00%"
      
      With ws.Columns(11).FormatConditions _
      .Add(xlCellValue, xlGreater, "=0")
      With .Interior
      .ColorIndex = 4
      End With
      End With
      
      With ws.Columns(11).FormatConditions _
      .Add(xlCellValue, xlLess, "=0")
      With .Interior
      .ColorIndex = 3
      End With
      End With

      ' Add headers for the Ticker and Value
      ws.Cells(1, 16).Value = "Ticker"
      ws.Cells(1, 17).Value = "Value"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      
      ' Determine greatest % of increase and decrease, and total volume
      increase_number = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
      decrease_number = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
      volume_number = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
          
      ' Print value for greatest % of increase and decrease, and total volume
      ws.Range("Q2") = increase_number
      ws.Range("Q2").NumberFormat = "0.00%"
      ws.Range("Q3") = decrease_number
      ws.Range("Q3").NumberFormat = "0.00%"
      ws.Range("Q4") = volume_number
  
      ' Find ticker symbol for  total, greatest % of increase and decrease, and average
      ws.Range("P2") = WorksheetFunction.Index(Range("I2:I" & LastRow), WorksheetFunction.Match(Range("Q2").Value, Range("K2:K" & LastRow), 0))
      ws.Range("P3") = WorksheetFunction.Index(Range("I2:I" & LastRow), WorksheetFunction.Match(Range("Q3").Value, Range("K2:K" & LastRow), 0))
      ws.Range("P4") = WorksheetFunction.Index(Range("I2:I" & LastRow), WorksheetFunction.Match(Range("Q4").Value, Range("L2:L" & LastRow), 0))

  Next ws

End Sub

