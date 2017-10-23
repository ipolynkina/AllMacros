' Заполнение шаблона для SAP HR
'
Private Sub fillPatternForSAP()
Call copySelectedData

Worksheets("copy").Select
Dim row, col As Integer
Dim colForAllDays As Integer
Const colForOneDay As Integer = 5
colForAllDays = ALL_DAYS * (HEADER_ROW * colForOneDay)

Worksheets("myCode").Select
For col = 4 To colForAllDays Step colForOneDay
    For row = 2 To ALL_GRAPHS_AND_ONE_HEADER Step 1
        Cells(row, col).Select
        ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,copy!C1:C2,2,0)"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Next row
    Worksheets("copy").Select
    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.delete Shift:=xlToLeft
    Worksheets("myCode").Select
Next col

Call copySelectedData

End Sub