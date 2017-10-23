' macro for remove excess nomenclature from the orders
' author: Polynkina Irina
' contact: irina.polynkina.dev@yandex.ru
' version: 1.0.0
' release: 29.04.2017
'
Const SHEET_NAME = "Переоформленные товары"
Const ID_ROW_FILTER = 2
Const ID_COL_NOM_KC = 22
Const ID_COL_NOM_BW = 37

Const BEGIN_BLOCK = "A"
Const END_BLOCK = "AP"
Const BEGIN_BLOCK_BW = "AF"
Const END_BLOCK_BW = "AP"

Const SIGN_COL_NOM_KC = "V"
Const SIGN_COL_ORDER_KC = "P"
Const SIGN_COL_NOM_BW = "AK"
Const SIGN_COL_ORDER_BW = "AF"
'
Sub rollUpData()

    ThisWorkbook.Save
    
    Call unFilters(SHEET_NAME)
    Call filterByColumn(SHEET_NAME, BEGIN_BLOCK, END_BLOCK, ID_ROW_FILTER, SIGN_COL_NOM_KC)
    Call filterByColumn(SHEET_NAME, BEGIN_BLOCK, END_BLOCK, ID_ROW_FILTER, SIGN_COL_ORDER_KC)
    Call filterByColumn(SHEET_NAME, BEGIN_BLOCK_BW, END_BLOCK_BW, ID_ROW_FILTER, SIGN_COL_NOM_BW)
    Call filterByColumn(SHEET_NAME, BEGIN_BLOCK_BW, END_BLOCK_BW, ID_ROW_FILTER, SIGN_COL_ORDER_BW)
    Call removeExcessNom(BEGIN_BLOCK_BW, END_BLOCK_BW, ID_ROW_FILTER, ID_COL_NOM_BW, ID_COL_NOM_KC)
    
    ThisWorkbook.Save
    MsgBox "Все готово! Жми OK!", vbExclamation, "version: 1.0.0"

End Sub
'
Private Sub unFilters(sheetName As String)

    On Error Resume Next
    If ActiveSheet.AutoFilterMode = True Then
        ActiveSheet.AutoFilterMode = False
    End If
    On Error GoTo 0

End Sub
'
Private Sub filterByColumn(sheetName As String, beginBlock As String, endBlock As String, _
                           numberRow As Integer, blockForFilter As String)
    
    Range(beginBlock + CStr(numberRow) + ":" + endBlock + CStr(numberRow)).Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets(sheetName).AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(sheetName).AutoFilter.Sort.SortFields. _
        add Key:=Range(blockForFilter + CStr(numberRow)), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(sheetName).AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
End Sub
'
Private Sub removeExcessNom(beginBlock As String, endBlock As String, rowForFilter As Integer, _
                            idNomBW As Integer, idNomKC As Integer)

    Dim row As Integer, rowForDelete As Integer
    row = rowForFilter + 1
    
    Do While Cells(row, idNomBW) <> ""
        If Cells(row, idNomBW) <> Cells(row, idNomKC) Then
            Range(beginBlock + CStr(row) + ":" + endBlock + CStr(row)).Select
            Selection.Delete Shift:=xlUp
            row = row - 1
        End If
        row = row + 1
    Loop

End Sub