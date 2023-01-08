Attribute VB_Name = "Module1"
'Create a script that loops through all the stocks for one year and outputs the following information:
  'The ticker symbol.
  'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  'The total stock volume of the stock.
  'Use conditional formatting that will highlight positive change in green and negative change in red.
'*********************************************************************************************************
'
Sub Stock_Analysis()

'Loop through all sheets
For Each ws In Worksheets

  ' Declare variables
  'to count last row in each worksheet in Column A
      Dim last_row As Long
      last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

  'Set variable to keep track of the location for each ticker
    Dim summary_table_row As Long
      summary_table_row = 2

  'Set variable for ticker name
      Dim ticker As String

  'Set variable for opening year stock
      Dim start_year As Double
      start_year = 0

  'Set variable for closing year stock
      Dim end_year As Double
      end_year = 0

      Dim yearly_change As Double
      Dim percent_change As Double

  'Set variable for holding a total count on the total volume of trade
      Dim total_stock_vol As Double
      total_stock_vol = 0

  'Set variable for counting number of entries for a particular ticker
      Dim ticker_count As Long
      ticker_count = 0

  'Set variables for greatest % increase, % decrease, and greatst stock volume
    Dim increase_ticker As String
    Dim increase_value As Double
    increase_value = 0

    Dim decrease_ticker As String
    Dim decrease_value As Double
    decrease_value = 0

    Dim total_vol_ticker As String
    Dim total_vol_value As Double
    total_vol_value = 0

    

    'Add column headers for ticker, yearly change, percent change, and volume
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"

    'Add column headers for greatest increase & decrease and greatest volume
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"

    'Formatting columns
      ws.Range("K:K").NumberFormat = "0.00%"
      ws.Range("Q2:Q3").NumberFormat = "0.00%"
      ws.Range("L:L").NumberFormat = "0"
      ws.Range("Q4").NumberFormat = "0"
      

    'row loop
    For r = 2 To last_row

      'Check if we are still within the same ticker name, if it is not....
      If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then

        'Set the ticker value
        ticker = ws.Cells(r, 1).Value

        'Print the Ticker in the summary table
         ws.Range("I" & summary_table_row).Value = ticker

        'Set the start_year value
         start_year = ws.Cells(r - ticker_count, 3).Value

        'Set the end_year value
         end_year = ws.Cells(r, 6).Value

        'Set the yearly_change value
         yearly_change = (end_year - start_year)

        'print the yearly change in the Summary Table
         ws.Range("J" & summary_table_row).Value = yearly_change

        'Calculating the percent change from opening price at the beginning of a given year
        'to the closing price at the end of that year.

        If (start_year <> 0) Then
          ws.Cells(summary_table_row, 11).Value = yearly_change / start_year
        Else
          
        End If
    
      
        'Add to the total stock volume
        total_stock_vol = total_stock_vol + ws.Cells(r, 7).Value

        'print the total stock volume in the Summary Table
        ws.Range("L" & summary_table_row).Value = total_stock_vol

        'Add one to the summary table row
        summary_table_row = summary_table_row + 1

        'Reset total stock volume
        total_stock_vol = 0

        'Reset ticker count
        ticker_count = 0


        'If the cell immediately following a row is the same ticker value...

        Else
          'Add current stock volume to previous stock volume
          total_stock_vol = total_stock_vol + ws.Cells(r, 7).Value
      
          'Add one to the ticker count
          ticker_count = ticker_count + 1
      End If
 
    Next r

    'conditional formatting that will highlight positive change in green and negative change in red.
    Dim con_last_row As Long
    con_last_row = ws.Cells(Rows.Count, 10).End(xlUp).Row

    For c = 2 To con_last_row

      If ws.Cells(c, 10).Value > 0 Then
          ws.Cells(c, 10).Interior.ColorIndex = 4
      ElseIf ws.Cells(c, 10).Value < 0 Then
          ws.Cells(c, 10).Interior.ColorIndex = 3
      End If

    Next c

    'for bonus table

    'Find the last non-blank cell in column I
    Dim last As Long
    last = ws.Cells(Rows.Count, 9).End(xlUp).Row

    'looping through
    For i = 2 To last
        
      'If current cell value is greater than the previous increase value...
      If ws.Cells(i, 11).Value > increase_value Then

        'set the new value
        increase_ticker = ws.Cells(i, 9).Value
        increase_value = ws.Cells(i, 11).Value

       'If current cell value is less than the previous decrease value...
       ElseIf ws.Cells(i, 11).Value < decrease_value Then

        'set the new value
        decrease_ticker = ws.Cells(i, 9).Value
        decrease_value = ws.Cells(i, 11).Value
                    
      End If
        
    Next i

    'Looping through the greatest stock volume
    For v = 2 To last
            
      'If current cell value is greater than previous Greatest Stock Volume...
      If ws.Cells(v, 12).Value > total_vol_value Then

        'set the new value
        total_vol_ticker = ws.Cells(v, 9).Value
        total_vol_value = ws.Cells(v, 12).Value
          
      End If
                         
    Next v

    'Assigning value with the newly defined %increase, %decrease, and greatest stock volume values
    ws.Range("P2").Value = increase_ticker
    ws.Range("Q2").Value = increase_value
        
    ws.Range("P3").Value = decrease_ticker
    ws.Range("Q3").Value = decrease_value
        
    ws.Range("P4").Value = total_vol_ticker
    ws.Range("Q4").Value = total_vol_value

    'Autofit entire column
    ws.Range("I:Q").EntireColumn.AutoFit

    'number format for the bonus table
    ws.Cells(4, 17).NumberFormat = "0.00E+00"

  Next ws

End Sub


