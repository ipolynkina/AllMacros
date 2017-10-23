' macro created for generate a report VOP min
' author: Polynkina Irina
' contact: irina.polynkina.dev@yandex.ru
' version: 1.0.0
' release: 17.12.2016
'
Public Const BEGINNING_OF_DATA = 2
Public Const NAME_NEW_LIST = "copy"
Public Const NAME_BASE_LIST = "VOP"
Public Const SIGN_SELLER = "Продавец"

Public Const ID_SHOP = 1
Public Const ID_REGION = 2
Public Const ID_DIVISION = 3
Public Const ID_PERS_NUM = 4
Public Const ID_FILENAME = 5
Public Const ID_POSITION = 6
Public Const ID_VOP = 10
Public Const ID_MONTH = 11
'
Sub createReportVOPMin()

    ThisWorkbook.Save
    Sheets.Add.name = NAME_NEW_LIST
    
    Call cleanOldData
    Call generateReportOnStaff(True)
    Call generateReportOnRegions(True)
    Call generateReportOnDivisions(True)
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(NAME_NEW_LIST).delete
    Sheets(NAME_BASE_LIST).Select
    Range("A1").Select
    ThisWorkbook.Save
    
    MsgBox "Все готово! Жми OK!", vbExclamation, "version: 1.1.0"
    
End Sub
'
Private Sub cleanOldData()

    Sheets(ReportOnStaff.RS_NAME_LIST).Select
    Rows(CStr(ReportOnStaff.RS_BEGINNING_OF_DATA) + ":" + CStr(ReportOnStaff.RS_BEGINNING_OF_DATA)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.delete Shift:=xlUp

    Sheets(ReportOnRegion.RR_NAME_LIST).Select
    Rows(CStr(ReportOnRegion.RR_BEGINNING_OF_DATA) + ":" + CStr(ReportOnRegion.RR_BEGINNING_OF_DATA)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.delete Shift:=xlUp
    
    Sheets(ReportOnDivision.RD_NAME_LIST).Select
    Rows(CStr(ReportOnDivision.RD_BEGINNING_OF_DATA) + ":" + CStr(ReportOnDivision.RD_BEGINNING_OF_DATA)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.delete Shift:=xlUp

End Sub
'
Public Sub copyOriginalData(isOpenFromOutside As Boolean)

    Sheets(NAME_BASE_LIST).Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(NAME_NEW_LIST).Select
    Range("A1").Select
    ActiveSheet.Paste

End Sub
'
Public Sub deleteRow(row As Integer)

    Rows(CStr(row) + ":" + CStr(row)).Select
    Selection.delete Shift:=xlUp

End Sub
'
Public Sub setRegions(sheetName As String)

    Sheets(sheetName).Select
    Dim row As Integer, endRow As Integer
    endRow = Cells(Rows.Count, 1).End(xlUp).row
    
    For row = BEGINNING_OF_DATA To endRow Step 1
        Cells(row, ID_REGION) = "=VLOOKUP(RC[-1],lib!C[-1]:C[1],3,0)"
    Next row
    
    Columns("B:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Sub
'
Public Sub setDivisions(sheetName As String)

    Sheets(sheetName).Select
    Dim row As Integer, endRow As Integer
    
    endRow = Cells(Rows.Count, 1).End(xlUp).row
    For row = BEGINNING_OF_DATA To endRow Step 1
        Cells(row, ID_DIVISION).Value = "=VLOOKUP(RC[-2],lib!C[-2]:C[-1],2,0)"
    Next row
    
    Columns("C:C").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Sub
'
Public Sub setFilterOnColumn(sheetName As String, indexRow As Integer, signBlock As String)
    
    Sheets(sheetName).Select
    Rows(CStr(indexRow) + ":" + CStr(indexRow)).Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets(sheetName).AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(sheetName).AutoFilter.Sort.SortFields.Add Key:= _
        Range(signBlock + CStr(indexRow)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
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
Public Sub setColor(rowForWrite As Integer, codeColor As Long, startBlock As String, endBlock As String)

    Set Rng = Range(startBlock + CStr(rowForWrite) + ":" + endBlock + CStr(rowForWrite))
    Rng.Interior.Color = codeColor
    
End Sub