' Если не находим график на листе directory_d - значит, это ночная смена; ставим -1.
' Если график на этом листе найден - ставим количество часов. Для листа directory_n все аналогично.
'
Private Sub searchSchedules()
Worksheets("selection").Select

Dim row As Integer
Const col As Integer = 78

For row = 2 To ALL_GRAPHS_AND_ONE_HEADER Step 1
    ' ищем на листе directory_d
    Cells(row, col + 11).FormulaR1C1 = "=VLOOKUP(RC[-45],directory_d!C[-88]:C[-87],2,0)"
    Cells(row, col + 10).FormulaR1C1 = "=VLOOKUP(RC[-46],directory_d!C[-87]:C[-86],2,0)"
    Cells(row, col + 9).FormulaR1C1 = "=VLOOKUP(RC[-47],directory_d!C[-86]:C[-85],2,0)"
    Cells(row, col + 8).FormulaR1C1 = "=VLOOKUP(RC[-48],directory_d!C[-85]:C[-84],2,0)"
    Cells(row, col + 7).FormulaR1C1 = "=VLOOKUP(RC[-49],directory_d!C[-84]:C[-83],2,0)"
    Cells(row, col + 6).FormulaR1C1 = "=VLOOKUP(RC[-50],directory_d!C[-83]:C[-82],2,0)"
    ' ищем на листе directory_n
    Cells(row, col + 5).FormulaR1C1 = "=VLOOKUP(RC[-39],directory_n!C[-82]:C[-81],2,0)"
    Cells(row, col + 4).FormulaR1C1 = "=VLOOKUP(RC[-40],directory_n!C[-81]:C[-80],2,0)"
    Cells(row, col + 3).FormulaR1C1 = "=VLOOKUP(RC[-41],directory_n!C[-80]:C[-79],2,0)"
    Cells(row, col + 2).FormulaR1C1 = "=VLOOKUP(RC[-42],directory_n!C[-79]:C[-78],2,0)"
    Cells(row, col + 1).FormulaR1C1 = "=VLOOKUP(RC[-43],directory_n!C[-78]:C[-77],2,0)"
    Cells(row, col + 0).FormulaR1C1 = "=VLOOKUP(RC[-44],directory_n!C[-77]:C[-76],2,0)"
Next row
      
Columns("BZ:CK").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Columns("BZ:CK").Select
Selection.Replace What:="#N/A", Replacement:="-1", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

End Sub