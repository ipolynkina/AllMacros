' Данные пользователя менять нельзя. Чтобы с ними работать - копируем на другой лист
'
Private Sub copyUserInput()

Worksheets("forUser").Select
Range("A3:AS3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("selection").Select
Range("A2").Select
ActiveSheet.Paste

End Sub