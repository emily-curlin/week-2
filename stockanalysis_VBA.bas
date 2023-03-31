Attribute VB_Name = "Module2"
Sub stock_analysis()

'define variables'
Dim ws As Worksheet
Dim row As Long
Dim new_row As Integer
Dim total_stock_volume As Double
Dim row_count As Long
Dim opening_price As Double
Dim closing_price As Double
Dim greatest_percent_increase_value As Double
Dim greatest_percent_decrease_value As Double
Dim greatest_total_volume_value As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease_ticker As String
Dim greatest_total_volume_ticker As String
Dim percent_change As Double

'apply to multiple sheets'
'Set ws = ActiveSheet'
For Each ws In Sheets

'starting values'
total_stock_volume = 0
new_row = 2
greatest_percent_increase_value = 0
greatest_percent_decrease_value = 0
greatest_total_volume_value = 0
greatest_percent_increase_ticker = ""
greatest_percent_decrease_ticker = ""
greatest_total_volume_ticker = ""



'loop through all rows in given worksheet, starting at row 2'
row_count = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
For row = 2 To row_count

'summary table headings'
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

'summary table headings part 2'
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"


'at first line in block of tickers, set stock volume equal to 0'
If ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
total_stock_volume = 0
opening_price = ws.Cells(row, 3).Value
End If

'add additional value to 0 for each value in block'
total_stock_volume = total_stock_volume + ws.Cells(row, 7).Value

'at end of block of tickers, store values in specified cells using specified variables'
If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
closing_price = ws.Cells(row, 6).Value
ws.Cells(new_row, 9).Value = ws.Cells(row, 1).Value
ws.Cells(new_row, 10).Value = closing_price - opening_price
percent_change = (closing_price - opening_price) / opening_price
ws.Cells(new_row, 11).Value = percent_change
ws.Cells(new_row, 12).Value = total_stock_volume
'insert color formating'
If (opening_price <= closing_price) Then
ws.Cells(new_row, 10).Interior.ColorIndex = 4
ws.Cells(new_row, 11).Interior.ColorIndex = 4
Else
ws.Cells(new_row, 10).Interior.ColorIndex = 3
ws.Cells(new_row, 11).Interior.ColorIndex = 3
End If

If total_stock_volume > greatest_total_volume_value Then
greatest_total_volume_value = total_stock_volume
greatest_total_volume_ticker = ws.Cells(row, 1).Value
End If

If percent_change > greatest_percent_increase_value Then
greatest_percent_increase_value = percent_change
greatest_percent_increase_ticker = ws.Cells(row, 1).Value
End If

If percent_change < greatest_percent_decrease_value Then
greatest_percent_decrease_value = percent_change
greatest_percent_decrease_ticker = ws.Cells(row, 1).Value
End If

'formating column 10 and 11'
ws.Cells(new_row, 10).NumberFormat = "$#,##0.00"
ws.Cells(new_row, 11).NumberFormat = "0.00%"
new_row = new_row + 1
End If


Next row

ws.Cells(2, 16).Value = greatest_percent_increase_ticker
ws.Cells(2, 17).Value = greatest_percent_increase_value
ws.Cells(3, 16).Value = greatest_percent_decrease_ticker
ws.Cells(3, 17).Value = greatest_percent_decrease_value
ws.Cells(4, 16).Value = greatest_total_volume_ticker
ws.Cells(4, 17).Value = greatest_total_volume_value

Next ws

End Sub


