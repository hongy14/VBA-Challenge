Attribute VB_Name = "Module1"
Sub AlphaTesting():

For Each ws In Worksheets


Dim Ticker As String
Dim Yearly_change As Double
Dim Percent_change As Double
Dim First_ticker As Boolean
Dim Open_price As Double
Dim Close_price As Double


Dim summary_table_row As Integer
summary_table_row = 2

Dim Total_stock_vol As Double
Total_stock_vol = 0

ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percentage Change"
ws.Cells(1, 12) = "Total Stock Volume"


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow
        
        If ws.Cells(i, 3) = 0 Then
        
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                Ticker = ws.Cells(i, 1)
            End If
        
        ElseIf ws.Cells(i + 1, 1) = ws.Cells(i, 1) Then
            Total_stock_vol = Total_stock_vol + ws.Cells(i, 7)
            
            If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
                Open_price = ws.Cells(i, 3)
            End If
            
        Else
            Ticker = ws.Cells(i, 1)
            Close_price = ws.Cells(i, 6)
            Yearly_change = Close_price - Open_price
            Percent_change = (Yearly_change / Open_price)
            Total_stock_vol = Total_stock_vol + ws.Cells(i, 7)
            ws.Cells(summary_table_row, 9) = Ticker
            ws.Cells(summary_table_row, 10) = Yearly_change
            
                If Yearly_change >= 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                ElseIf Yearly_change < 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                End If
            
            ws.Cells(summary_table_row, 11) = Percent_change
            ws.Range("K" & summary_table_row) = Format(ws.Range("K" & summary_table_row), "Percent")
            ws.Cells(summary_table_row, 12) = Total_stock_vol
            summary_table_row = summary_table_row + 1
            Yearly_change = 0
            Percentage_change = 0
            Total_stock_vol = 0
            
           
        End If
        
            
    Next i


Next ws


End Sub

