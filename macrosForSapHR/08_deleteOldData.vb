' Удаляем старые данные
'
Private Sub deleteOldData()

Worksheets("myCode").Select
Range("D2:FB2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
ActiveWorkbook.Save

End Sub