Attribute VB_Name = "Module1"
Sub ShowMeTheMoney()
'Create a script that will loop through all the stocks for one year and output the following information:

'Declare variables:
Dim ws As Worksheet
Dim lastrow As Long
Dim Tickers As String
Dim year_open As Double
Dim year_close As Double
Dim volume As LongLong
Dim yearly_change As Double
Dim percent_change As Double
Dim output_table_row As Integer

'Begin running through your worksheet and set up framework for the script to run on all worksheets
'Headers for Output Columns

For Each ws In ThisWorkbook.Worksheets

    ws.Cells(1, 10).Value = "Ticker" 'Set up output column headers
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Trading Volume"

'set your starting point for the loop
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
output_table_row = 2

    For i = 2 To lastrow 'Begin loop
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then 'Find very firstiteration of a particular ticker
        year_open = ws.Cells(i, 3).Value ' Grab this value as it is the first trading day price of the year
        End If
        
'The total stock volume of the stock
        volume = volume + ws.Cells(i, 7) 'Sum volume as it goes through the rows of the same ticker
        
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then 'If the value of the ticker cell does not equal the one below it then...
'The ticker symbol
            ticker = ws.Cells(i, 1).Value 'Grab ticker symbol
            ws.Cells(output_table_row, 10).Value = ticker 'Output Ticker
            ws.Cells(output_table_row, 13).Value = volume 'Output volume
            year_close = ws.Cells(i, 6) 'Grab year close price
            
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
            yearly_change = (year_close - year_open) 'Grab YTD nominal change
            ws.Cells(output_table_row, 11).Value = yearly_change 'Output yearly change
            
'You should also have conditional formatting that will highlight positive change in green and negative change in red.
            If yearly_change > 0 Then
                ws.Cells(output_table_row, 11).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                ws.Cells(output_table_row, 11).Interior.ColorIndex = 3
            End If
            
'Percent change from opening price at the beginning of a given year to the closing price ...
            If year_open = 0 And year_close = 0 Then
                percent_change = 0
                ws.Cells(output_table_row, 12).Value = percent_change 'output percent change
                ws.Cells(output_table_row, 12).NumberFormat = "0.00%" 'change to percentage format
   
            ElseIf year_open = 0 And year_close > 0 Then 'Can't include new listings in this because growth is technically infinite when you start at 0 and end anywhere. And excel will show this as 100%, which is also inaccurate.
                Dim percent_change_newlisting As String
                percent_change_newlisting = "New Listing"
                ws.Cells(output_table_row, 12).Value = percent_change_newlisting
            'WHY DOESNT THIS WORK THOUGH - See PLNT in 2015. Year open is 0 and then sometime mid-year it begins trading at a price. However, this ticker is not in my summary table...
            
            Else
                percent_change = (yearly_change / year_close) 'Grab YTD % change
                ws.Cells(output_table_row, 12).Value = percent_change ' output percent change
                ws.Cells(output_table_row, 12).NumberFormat = "0.00%" 'change to percentage format
        
        output_table_row = output_table_row + 1 'Create next Output Row for next Ticker info
            End If
        'Reset volume as you finish gathering data on all trading days of each unique ticker - nothing else needs to reset
        volume = 0
        
        End If
        
    Next i
    
'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
'Start with creating output area

ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"

'Declare variables for each category
Dim best_ticker As String
Dim best_value As Double

Dim worst_ticker As String
Dim worst_value As Double

Dim highest_volume_ticker As String
Dim highest_volume_value As Double

best_value = ws.Cells(2, 12).Value
worst_value = ws.Cells(2, 12).Value
highest_volume_value = ws.Cells(2, 13).Value

'Tell VBA where last data set is
lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

'For best performer find highest percent change among summary output.
For j = 2 To lastrow
    If ws.Cells(j, 12).Value > best_value Then
        best_value = ws.Cells(j, 12).Value
        best_ticker = ws.Cells(j, 10).Value
    End If
    
'For worst performer, lowest percent change
    If ws.Cells(j, 12).Value < worst_value Then
        worst_value = ws.Cells(j, 12).Value
        worst_ticker = ws.Cells(j, 10).Value
    End If
'Highest trading volume
    If ws.Cells(j, 13).Value > highest_volume_ticker Then
        highest_volume_value = ws.Cells(j, 13).Value
        highest_volume_ticker = ws.Cells(j, 10).Value
    End If
        
Next j

'Last step is to import output values in the table you made at the beginning of this for loop
ws.Cells(2, 17).Value = best_ticker
ws.Cells(2, 18).Value = best_value
ws.Cells(2, 18).NumberFormat = "0.00%"
ws.Cells(3, 17).Value = worst_ticker
ws.Cells(3, 18).Value = worst_value
ws.Cells(3, 18).NumberFormat = "0.00%"
ws.Cells(4, 17).Value = highest_volume_ticker
ws.Cells(4, 18).Value = highest_volume_value

'Autofit table columns for desired prettiness
ws.Columns("J:R").EntireColumn.AutoFit

'Repeat the entire thing on the next worksheet
Next ws

End Sub



