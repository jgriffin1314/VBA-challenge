Attribute VB_Name = "Module1"
Sub multi_yr_stock_data()

Dim ticker As String
Dim opening_price As Double
Dim closing_price As Double
Dim volume As Double
Dim summary_table_row As Integer

summary_table_row = 2

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Volume"
Cells(1, 11).Value = "Opening Price for Year"
Cells(1, 12).Value = "Closing Price for Year"
Cells(1, 13).Value = "Yearly Change"
Cells(1, 14).Value = "Percent Change"

Range("K2") = Cells(2, 3).Value


For i = 2 To 100000

If Cells(i + 1, 1).Value <> Cells(i, 1) Then

ticker = Cells(i, 1).Value
opening_price = Cells(i + 1, 3).Value

volume = volume + Cells(i, 7)

closing_price = Cells(i, 6)


Range("I" & summary_table_row).Value = ticker
Range("J" & summary_table_row).Value = volume
Range("K" & summary_table_row + 1).Value = opening_price
Range("L" & summary_table_row).Value = closing_price
Range("M" & summary_table_row).Value = Cells(summary_table_row, 11).Value - Cells(summary_table_row, 12).Value
Range("N" & summary_table_row).Value = (Cells(summary_table_row, 11).Value - Cells(summary_table_row, 12).Value) / Cells(summary_table_row, 11).Value

If Range("M" & summary_table_row).Value > 0 Then
Range("M" & summary_table_row).Interior.ColorIndex = 4
Else
Range("M" & summary_table_row).Interior.ColorIndex = 3
End If

summary_table_row = summary_table_row + 1

volume = 0

Else

volume = volume + Cells(i, 7)


End If

Next i


End Sub
