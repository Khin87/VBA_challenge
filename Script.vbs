
Sub ticker_sample()

Dim ticker_name As String
Dim ticker_vol_total As Double
ticker_vol_total = 0
Dim summary_table_row As Integer
summary_table_row = 2

Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'set names
Cells(1, 9).Value = "ticker_name"
Cells(1, 12).Value = "ticker_vol_total"

'loop through all stocks
For i = 2 To lastrow

'if the same stock , if not
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'set ticker name
ticker_name = Cells(i, 1).Value

'add ticker vol total
ticker_volume_total = ticker_volume_total + Cells(1, 7).Value

'heading in summary table

Range("I" & summary_table_row).Value = ticker_name
Range("L" & summary_table_row).Value = ticker_vol_total

'add summary table row
summary_table_row = summary_table_row + 1

'reset ticker vol total
ticker_vol_total = 0

'if cell immediate following row is same

Else
'add ticker vol toal
ticker_vol_total = ticker_vol_total + Cells(i, 7).Value
End If
Next i



End Sub

Sub yearlychange()

'define

Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim summary_table_row As Integer
summary_table_row = 2

'set name
Cells(1, 10).Value = "Yearly_Change"
Cells(1, 11).Value = "Percent_Change"

Dim lastrow As Long
lastrow = Cells.SpecialCells(xlCellTypeLastCell).Row

'create loop
For i = 2 To lastrow
Cells(i, 10) = Cells(i, 6) - Cells(i, 3)

'percent change
Cells(i, 11) = (1 - (Cells(i, 6) / Cells(i, 3))) * 100
Next i


End Sub


Sub color():


Dim i As Double
For i = 2 To lastrow
lastrow = Cells.SpecialCells(xlCellTypeLastCell).Row

      Set Cell = Range("J" & i)
     
      If Cells.Value >= 0 Then
      
      'set cellvalue
      Range("J", i).Value = positive
      Range("J", i).Value = Interior.ColorIndex = 3
      
      Else: Range("J", i).Value = positive
      Range("J", i).Value = Interior.ColorIndex = 4
      End If
      
     
   Next i

End Sub


