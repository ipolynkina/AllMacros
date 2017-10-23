' Копируем подобранные данные на новый лист. Изменять нельзя, они понадобятся дважды.
'
Private Sub copySelectedData()

Worksheets("selection").Select
Range("AT1:BY1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Selection.Copy
Worksheets("copy").Select
Range("A1").Select
ActiveSheet.Paste

End Sub