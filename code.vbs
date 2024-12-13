Sub Qtly_Stock_review()

  'Define variables for
  ' 1. ticker
  ' 2. qt_chg, qt_open, qt_close
  ' 3. perc_chg
  ' 4. tot_stock_vol

  Dim ticker As String
  Dim qt_change As Double, qt_open As Double, qt_close As Double
  Dim pct_change As Double
  Dim tot_stock_vol As Double

  'variables to track rows
  Dim summary_row_num As Long, last_row As Long, ticker_change As Long, curr_row_num As Long

  'var to iterate over all worksheets
  Dim ws As Worksheet

  ' Variables for tracking greatest % increase, % decrease, and total volume
  Dim max_pct_increase As Double, min_pct_decrease As Double, max_tot_vol As Double
  Dim max_pct_inc_ticker As String, min_pct_dec_ticker As String, max_tot_vol_ticker As String

  'Iterate over worksheets
  For Each ws In Worksheets

    ' Initialize tracking variables
    max_pct_increase = 0
    min_pct_decrease = 0
    max_tot_vol = 0

    'assign summary col headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'initalise summary_row_num
    summary_row_num = 2

    'init the ticker_change tracker row to first data row
    ticker_change = 2

    'Find the last row with Tickers in Col A
    last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

      'For each ticker summarize and loop the quarterly change, percent change, and total stock volume
      For curr_row_num = 2 To last_row

        'add tot_stock_vol to running total
        tot_stock_vol = tot_stock_vol + ws.Cells(curr_row_num, 7).Value

        'chk for ticker change - if changed, execute summary stats
        If ws.Cells(curr_row_num + 1, 1).Value <> ws.Cells(curr_row_num, 1).Value Then

        'Get ticker value for records
        ticker = ws.Cells(curr_row_num, 1).Value

        ' Get qt_open from col "C" and qt_close from col 6
        qt_open = ws.Cells(ticker_change, 3).Value
        qt_close = ws.Cells(curr_row_num, 6).Value

        'Calculate pct_change
        If qt_open = 0 Then
          pct_change = qt_close
        Else
          qt_change = qt_close - qt_open
          pct_change = qt_change / qt_open
        End If

        'assign summary values
        ws.Cells(summary_row_num, 9).Value = ticker

        'assign the qt_change
        ws.Cells(summary_row_num, 10).Value = qt_change

        'assign green background color if +ve else red if -ve
        If qt_change > 0 Then
          ws.Cells(summary_row_num, 10).Interior.Color = vbGreen
        ElseIf qt_change < 0 Then
          ws.Cells(summary_row_num, 10).Interior.Color = vbRed
        End If

        'assign pct along with the number format
        ws.Cells(summary_row_num, 11).Value = pct_change
        ws.Cells(summary_row_num, 11).NumberFormat = "0.00%"

        'assign total stock volume
        ws.Cells(summary_row_num, 12).Value = tot_stock_vol

        ' Track greatest % increase
        If pct_change > max_pct_increase Then
            max_pct_increase = pct_change
            max_pct_inc_ticker = ticker
        End If

        ' Track greatest % decrease
        If pct_change < min_pct_decrease Then
            min_pct_decrease = pct_change
            min_pct_dec_ticker = ticker
        End If

        ' Track greatest total volume
        If tot_stock_vol > max_tot_vol Then
            max_tot_vol = tot_stock_vol
            max_tot_vol_ticker = ticker
        End If

        'Summary for one ticker completed, increment the summary row
        summary_row_num = summary_row_num + 1

        'reinit the summary vars
        tot_stock_vol = 0
        qt_change = 0
        pct_change = 0

        'Init the ticker change to curr_row_num +1
        ticker_change = curr_row_num + 1

      End If
      Next curr_row_num

      ' Display the results for greatest % increase, % decrease, and total volume
      ws.Range("O2").Value = "Greatest % increase"
      ws.Range("O3").Value = "Greatest % decrease"
      ws.Range("O4").Value = "Greatest total volume"

      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"

      ws.Range("P2").Value = max_pct_inc_ticker
      ws.Range("Q2").Value = max_pct_increase
      ws.Range("Q2").NumberFormat = "0.00%"

      ws.Range("P3").Value = min_pct_dec_ticker
      ws.Range("Q3").Value = min_pct_decrease
      ws.Range("Q3").NumberFormat = "0.00%"

      ws.Range("P4").Value = max_tot_vol_ticker
      ws.Range("Q4").Value = max_tot_vol
      ws.Range("Q4").NumberFormat = "0.00E+00"

      ' Auto-resize columns "I" to "Q"
      ws.Columns("I:Q").AutoFit

  Next ws
End Sub


