Sub stock_info_big()

On Error Resume Next
    

Dim row As Double
Dim total_vol As Double
Dim first_open As Double
Dim last_close As Double
Dim yearly_change As Double
Dim c As Integer
Dim i As Integer

Dim ws_count As Integer
ws_count = ActiveWorkbook.Worksheets.Count
For i = 1 To ws_count


Worksheets(i).Cells(1, 9).Value = "Ticker"
Worksheets(i).Cells(1, 10).Value = "Yearly Change"
Worksheets(i).Cells(1, 12).Value = "Percent Change"
Worksheets(i).Cells(1, 11).Value = "Total Stock Vol"

Worksheets(i).Cells(1, 16).Value = "Ticker"
Worksheets(i).Cells(1, 17).Value = "Value"
Worksheets(i).Cells(2, 15).Value = "Greatest % Increase"
Worksheets(i).Cells(3, 15).Value = "Greatest % Decrease"
Worksheets(i).Cells(4, 15).Value = "Greatest Total Volume"

row = 2
total_vol = 0
    'deterimines first_open, percent_change, and yearly_change
    For j = 2 To Worksheets(i).Cells(Rows.Count, 1).End(xlUp).row
        If Worksheets(i).Cells(j, 1).Value <> Worksheets(i).Cells(j + 1, 1).Value Then
            last_close = Worksheets(i).Cells(j, 6).Value
            Worksheets(i).Cells(row, 10).Value = last_close - first_open
            Worksheets(i).Cells(row, 12).Value = (last_close - first_open) / first_open
            first_open = 0
            last_close = 0
            'takes vol 1st time
            Worksheets(i).Cells(row, 9).Value = Worksheets(i).Cells(j, 1).Value
            total_vol = total_vol + Worksheets(i).Cells(j, 7).Value
            'takes ticker
            Worksheets(i).Cells(row, 11).Value = total_vol
            row = row + 1
            total_vol = 0
        ElseIf Worksheets(i).Cells(j - 1, 1).Value <> Worksheets(i).Cells(j, 1).Value Then
            first_open = Worksheets(i).Cells(j, 3).Value
            'takes vol 2nd time
            total_vol = total_vol
        Else
            'takes vol 3rd time
            total_vol = total_vol + Worksheets(i).Cells(j, 7).Value
        End If
   Next j

'finds max/min percentage and max volume
last_row_percentage = Cells(Rows.Count, 12).End(xlUp).row
last_row_volume = Cells(Rows.Count, 11).End(xlUp).row

    ActiveWorkbook.Worksheets(i).Cells(2, 17).Value = WorksheetFunction.Max(Worksheets(i).Range("L2:L" & ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count))
    ActiveWorkbook.Worksheets(i).Cells(3, 17).Value = WorksheetFunction.Min(Worksheets(i).Range("L2:L" & ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count))
    ActiveWorkbook.Worksheets(i).Cells(4, 17).Value = WorksheetFunction.Max(Worksheets(i).Range("k2:k" & ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count))
    
    'formats color
    For l = 2 To ActiveWorkbook.Worksheets(i).Range("j1").CurrentRegion.Rows.Count
            If ActiveWorkbook.Worksheets(i).Cells(l, 10).Value > 0 Then
                With ActiveWorkbook.Worksheets(i).Cells(l, 10).Interior
                    .ColorIndex = 4
                    .TintAndShade = 0.6
                End With
            Else
                With ActiveWorkbook.Worksheets(i).Cells(l, 10).Interior
                    .ColorIndex = 3
                    .TintAndShade = 0.6
                End With
            End If
        Next l

    'displays the ticker for max/min and max volume and formats color
    For j = 2 To Worksheets(i).Cells(Rows.Count, 9).End(xlUp).row
        If Worksheets(i).Cells(j, 12).Value = Worksheets(i).Cells(2, 17).Value Then
            Worksheets(i).Range("P2").Value = Worksheets(i).Cells(j, 9).Value
        ElseIf Worksheets(i).Cells(j, 12).Value = Worksheets(i).Cells(3, 17).Value Then
            Worksheets(i).Range("P3").Value = Worksheets(i).Cells(j, 9).Value
        ElseIf Worksheets(i).Cells(j, 11).Value = Worksheets(i).Cells(4, 17).Value Then
            Worksheets(i).Range("P4").Value = Worksheets(i).Cells(j, 9).Value
        End If
    Next j
    


Next i
End Sub














