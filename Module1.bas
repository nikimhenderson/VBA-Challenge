Attribute VB_Name = "Module1"
Sub Stock_Market_Value()

Dim ticker As String
Dim total_stock_volume As Double
Dim percent_change As Double
Dim i, j As Integer
Dim closing_price As Double
Dim yearly_change As Double
Dim starting_price As Double
Dim summary_table_row As Integer
Dim ws As Worksheet
Dim greatest_per_increase As Double
Dim greatest_per_decrease As Double
Dim greatest_total_vol As Double
Dim MaxIndex As Long
Dim MinIndex As Long
Dim VolIndex As Long

'Start worksheet loop
For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total_Stock_Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Initialize variables
last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
yearly_change = 0
summary_table_row = 2
total_stock_volume = 0

For i = 2 To last_row
    If i = 2 Then
        starting_price = ws.Cells(i, 3).Value
    End If
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value
        closing_price = ws.Cells(i, 6).Value
        yearly_change = closing_price - starting_price
        percent_change = yearly_change / starting_price
        
        ws.Range("I" & summary_table_row).Value = ticker
        ws.Range("J" & summary_table_row).Value = yearly_change
        ws.Range("K" & summary_table_row).Value = FormatPercent(percent_change)
        ws.Range("L" & summary_table_row).Value = total_stock_volume
        
        
        summary_table_row = summary_table_row + 1
        starting_price = ws.Cells(i + 1, 3).Value
        total_stock_volume = 0
    Else
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
    End If
    'format the output
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    

Next i

    'Greatest percent increase, decrease and total volume
    greatest_per_increase = WorksheetFunction.Max(ws.Range("K2:K" & last_row))
    greatest_per_decrease = WorksheetFunction.Min(ws.Range("K2:K" & last_row))
    greatest_total_vol = WorksheetFunction.Max(ws.Range("L2:L" & last_row))
    ws.Cells(2, 17).Value = "%" & greatest_per_increase * 100
    ws.Cells(3, 17).Value = "%" & greatest_per_decrease * 100
    ws.Cells(4, 17).Value = greatest_total_vol
    
    'Find the index
    MaxIndex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
    MinIndex = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
    VolIndex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & last_row)), ws.Range("L2:L" & last_row), 0)
    'Find the ticker based on the index
    ws.Cells(2, 16).Value = ws.Cells(MaxIndex + 1, 9).Value
    ws.Cells(3, 16).Value = ws.Cells(MinIndex + 1, 9).Value
    ws.Cells(4, 16).Value = ws.Cells(VolIndex + 1, 9).Value
    
    
Next ws


End Sub


