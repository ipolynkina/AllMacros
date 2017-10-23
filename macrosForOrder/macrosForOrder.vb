' макрос для автоматического составления приказа
' по всем вопросам писать на itcoder71@gmail.com
' version: 3.0.0
' release: 04.06.2017

' sheet DATA
Const SHEET_NAME_DATA = "ИО"

Const COL_NUMBER_SHOP_DATA = "A"
Const COL_PERS_NUM_DATA = "B"
Const COL_FULL_NAME_DATA = "C"
Const COL_SUM_DATA = "G"

' sheet ORDER
Const SHEET_NAME_ORDER = "!для приказа"

Const ROW_COUNTER = 1
Const COL_COUNTER = 14

Const BEGIN_RANGE_ORDER = "A"
Const END_RANGE_ORDER = "I"

Const COL_PERS_NUM_ORDER = "A"
Const COL_FULL_NAME_ORDER = "B"
Const COL_NUMBER_SHOP_ORDER = "C"
Const COL_SUM_WITHOUT_RKSN = "D"
Const COL_SUM_RK = "E"
Const COL_SUM_SN = "F"
Const COL_SUM_ORDER = "G"
Const COL_RK_ORDER = "H"
Const COL_SN_ORDER = "I"

Const AMOUNT_OF_HEADER = 2
Const INDEX_FIRST_ROW = 3
Const INDEX_SUM_WITHOUT_RKSN = 4
Const INDEX_SUM_RK = 5
Const INDEX_SUM_SN = 6
Const INDEX_RK = 8
Const INDEX_SN = 9

Sub macrosForOrder()

    ActiveWorkbook.Save
    
    ' читаем номер строки, с которой начать обработку
    Dim beginRow As String
    beginRow = Cells(ROW_COUNTER, COL_COUNTER)
    If beginRow = "" Or beginRow = "0" Then
        MsgBox "Необходимо ввести номер строки", vbCritical, "Ошибка"
        Exit Sub
    End If
    
    ' проверяем, что количество строк больше 0 (иначе обрабатывать нечего)
    Sheets(SHEET_NAME_DATA).Select
    Dim endRow As Long
    Dim allRow As Long
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    allRow = endRow - beginRow + 1
    If allRow < 1 Then
        MsgBox "В приказе не может быть " & allRow & " человек", vbCritical, "Ошибка"
        Exit Sub
    End If
    
    ' отключаем действия пользователя для ускорения
    Application.ScreenUpdating = False
    Application.Interactive = False
    Application.EnableCancelKey = xlDisabled
    
    ' удаляем старые данные
    Sheets(SHEET_NAME_ORDER).Select
    Range(BEGIN_RANGE_ORDER + CStr(INDEX_FIRST_ROW) + ":" + END_RANGE_ORDER + CStr(INDEX_FIRST_ROW)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ' переносим на другой лист БС, ТН, ФИО, сумму
    Call copyColumn(SHEET_NAME_DATA, beginRow, COL_NUMBER_SHOP_DATA, SHEET_NAME_ORDER, INDEX_FIRST_ROW, COL_NUMBER_SHOP_ORDER)
    Call copyColumn(SHEET_NAME_DATA, beginRow, COL_PERS_NUM_DATA, SHEET_NAME_ORDER, INDEX_FIRST_ROW, COL_PERS_NUM_ORDER)
    Call copyColumn(SHEET_NAME_DATA, beginRow, COL_FULL_NAME_DATA, SHEET_NAME_ORDER, INDEX_FIRST_ROW, COL_FULL_NAME_ORDER)
    Call copyColumn(SHEET_NAME_DATA, beginRow, COL_SUM_DATA, SHEET_NAME_ORDER, INDEX_FIRST_ROW, COL_SUM_ORDER)
    
    ' в БС заменяем русские буквы "а", "А" и английскую "a" на английскую "A" (т.к. РКСН будем искать по английской "A")
    Sheets(SHEET_NAME_ORDER).Select
    Columns(COL_NUMBER_SHOP_ORDER + ":" + COL_NUMBER_SHOP_ORDER).Select
    
    Selection.Replace What:="а", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="А", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="a", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    ' ищем значения РК и СН (в процентах)
    Range(COL_RK_ORDER + CStr(INDEX_FIRST_ROW)).Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],рксн!C[-7]:C[-4],4,0)"
    Range(COL_SN_ORDER + CStr(INDEX_FIRST_ROW)).Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],рксн!C[-8]:C[-4],5,0)"
    
    ' рассчитываме сумму премии, РК и СН (в рублях)
    Range(COL_SUM_WITHOUT_RKSN + CStr(INDEX_FIRST_ROW)).Select
    ActiveCell.FormulaR1C1 = "=RC[3]/(1+(RC[4]+RC[5])/100)"
    Range(COL_SUM_RK + CStr(INDEX_FIRST_ROW)).Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*(RC[3]/100)"
    Range(COL_SUM_SN + CStr(INDEX_FIRST_ROW)).Select
    ActiveCell.FormulaR1C1 = "=RC[-2]*(RC[3]/100)"
    
    ' если обрабатываем более одной строки - протягиваем формулы
    If allRow > 1 Then
        Range(COL_RK_ORDER + CStr(INDEX_FIRST_ROW)).Select
        Selection.AutoFill Destination:=Range(Selection, Cells(allRow + AMOUNT_OF_HEADER, INDEX_RK))
        Range(COL_SN_ORDER + CStr(INDEX_FIRST_ROW)).Select
        Selection.AutoFill Destination:=Range(Selection, Cells(allRow + AMOUNT_OF_HEADER, INDEX_SN))
        Range(COL_SUM_WITHOUT_RKSN + CStr(INDEX_FIRST_ROW)).Select
        Selection.AutoFill Destination:=Range(Selection, Cells(allRow + AMOUNT_OF_HEADER, INDEX_SUM_WITHOUT_RKSN))
        Range(COL_SUM_RK + CStr(INDEX_FIRST_ROW)).Select
        Selection.AutoFill Destination:=Range(Selection, Cells(allRow + AMOUNT_OF_HEADER, INDEX_SUM_RK))
        Range(COL_SUM_SN + CStr(INDEX_FIRST_ROW)).Select
        Selection.AutoFill Destination:=Range(Selection, Cells(allRow + AMOUNT_OF_HEADER, INDEX_SUM_SN))
    End If
    
    ' ищем дублирующиеся табельные
    Call findDuplicates(SHEET_NAME_DATA, beginRow, COL_PERS_NUM_DATA)

    ' сохраняем изменения
    Sheets(SHEET_NAME_DATA).Select
    Range(COL_NUMBER_SHOP_DATA + CStr(endRow)).Select
    Sheets(SHEET_NAME_ORDER).Select
    Range(BEGIN_RANGE_ORDER + CStr(1)).Select
    ActiveWorkbook.Save
    
    ' и разрешаем действия пользователя
    Application.ScreenUpdating = True
    Application.Interactive = True
    Application.EnableCancelKey = xlInterrupt

End Sub

Private Sub copyColumn(sheetNameOrigin As String, rowOrigin As String, colOrigin As String, _
                       sheetNameCopy As String, rowCopy As String, colCopy As String)

    Sheets(sheetNameOrigin).Select
    Range(colOrigin + rowOrigin).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets(sheetNameCopy).Select
    Range(colCopy + rowCopy).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Sub

Private Sub filterByColumn(sheetName As String, beginBlock As String, endBlock As String, _
                           numberRow As Integer, blockForFilter As String)

    Range(beginBlock + CStr(numberRow) + ":" + endBlock + CStr(numberRow)).Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets(sheetName).AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(sheetName).AutoFilter.Sort.SortFields. _
        Add Key:=Range(blockForFilter + CStr(numberRow)), SortOn:=xlSortOnValues, Order:=xlAscending, _
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

Private Sub findDuplicates(sheetNameOrigin As String, rowOrigin As String, colOrigin As String)

    Const SHEET_DUPLICATES = "duplicates"
    Const COL_PERS_NUM_DUPLICATES = "A"
    Const INDEX_PERS_NUM = 1

    Sheets.Add.Name = SHEET_DUPLICATES
    Call copyColumn(sheetNameOrigin, rowOrigin, colOrigin, SHEET_DUPLICATES, INDEX_PERS_NUM + 1, COL_PERS_NUM_DUPLICATES)

    Range("A1").Select
    ActiveCell.FormulaR1C1 = "PersNum"
    Call filterByColumn(SHEET_DUPLICATES, COL_PERS_NUM_DUPLICATES, COL_PERS_NUM_DUPLICATES, INDEX_PERS_NUM, COL_PERS_NUM_DUPLICATES)

    Dim row As Integer
    row = INDEX_PERS_NUM + 1

    Do While Cells(row, INDEX_PERS_NUM) <> ""
        If Cells(row, INDEX_PERS_NUM) = Cells(row - 1, INDEX_PERS_NUM) Then
            MsgBox "Табельный номеро повторяется: " & Cells(row, INDEX_PERS_NUM), vbExclamation, "version: 3.0.0"
        End If
        row = row + 1
    Loop

    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(SHEET_DUPLICATES).Delete

End Sub