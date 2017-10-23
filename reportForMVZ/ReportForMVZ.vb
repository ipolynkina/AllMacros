' макрос для автоматизации отчета по МВЗ
' по всем вопросам писать на irina.polynkina.dev@yandex.ru
' version: 2.0.0

Public Const SHEET_INITIAL_DATE = "исх"
Public Const SHEET_WITH_VACANCY = "обр_с_вак"
Public Const SHEET_WITHOUT_VACANCY = "обр_без_вак"
Public Const SHEET_CATALOG = "справочник"

'
Sub RUN()

    ThisWorkbook.Save
    
    Call deleteSheets
    Call makeBeauty
    Call removeColumns
    
    Call createAndHandleSheetWithVacancy
    Call createAndHandleSheetWithoutVacancy
    Call addRate(SHEET_WITHOUT_VACANCY)
    Call addSubdivision(SHEET_WITHOUT_VACANCY)
    
    Sheets(SHEET_INITIAL_DATE).Select
    Range("A1").Select
    
    Call alignColumns(SHEET_WITH_VACANCY)
    Sheets(SHEET_WITH_VACANCY).Select
    Range("A1").Select
    
    Call alignColumns(SHEET_WITHOUT_VACANCY)
    Sheets(SHEET_WITHOUT_VACANCY).Select
    Range("A1").Select
    
    Call plus_km_and_im
    
End Sub

'
Private Sub deleteSheets()

    Application.DisplayAlerts = False
    
    Dim i As Long
    For i = Sheets.Count To 1 Step - 1
        If Sheets(i).Name = SHEET_WITH_VACANCY _
        Or Sheets(i).Name = SHEET_WITHOUT_VACANCY Then
            Sheets(i).Delete
        End If
    Next
    
    Application.DisplayAlerts = True

End Sub

'
Private Sub makeBeauty()

    Sheets.Add.Name = SHEET_WITH_VACANCY
    Call copySheet(SHEET_INITIAL_DATE, SHEET_WITH_VACANCY)
    
    Sheets(SHEET_WITH_VACANCY).Select
    
    ' заливаем белым
    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' шрифт поменьше
    With Selection.Font
        .Name = "Calibri"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    ' закрепляем диапазон
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    ' рисуем рамки
    Call setBorder(xlEdgeLeft)
    Call setBorder(xlEdgeTop)
    Call setBorder(xlEdgeBottom)
    Call setBorder(xlEdgeRight)
    Call setBorder(xlInsideVertical)
    Call setBorder(xlInsideHorizontal)
    
    ' красим шапку цветом
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
End Sub

'
Private Sub setBorder(parameter As String)

    With Selection.Borders(parameter)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub

'
Private Sub removeColumns()

    Sheets(SHEET_WITH_VACANCY).Select
    
    Columns("E:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:N").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:M").Select
    Selection.Delete Shift:=xlToLeft
    
End Sub

'
Private Sub createAndHandleSheetWithVacancy()

    Const MVZ = 6
    Const FULL_MVZ = 9
    
    Const ORDER = 7
    Const FULL_ORDER = 10
    
    Const SPP = 8
    Const FULL_SPP = 11

    Sheets(SHEET_WITH_VACANCY).Select
    Columns("A:K").Select
    Selection.NumberFormat = "General"
    
    Range("A1").Select
    Dim endRow As Long
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' МВЗ
    For beginRow = 2 To endRow Step 1
        If Cells(beginRow, MVZ) = "" Then
            If Cells(beginRow, FULL_MVZ) = "" Then
                Cells(beginRow, FULL_MVZ) = "0"
            End If
            Cells(beginRow, MVZ) = Cells(beginRow, FULL_MVZ)
        End If
    Next beginRow
    
    ' заказ
    For beginRow = 2 To endRow Step 1
        If Cells(beginRow, ORDER) = "" Then
            Cells(beginRow, ORDER) = Cells(beginRow, FULL_ORDER)
        End If
    Next beginRow
    
    ' СПП-Элемент
    For beginRow = 2 To endRow Step 1
        If Cells(beginRow, FULL_SPP) <> "" Then
            Cells(beginRow, SPP) = Cells(beginRow, FULL_SPP)
        End If
    Next beginRow
    
    Columns("I:L").Select
    Selection.Delete Shift:=xlToLeft
    
End Sub

'
Private Sub createAndHandleSheetWithoutVacancy()

    Sheets.Add.Name = SHEET_WITHOUT_VACANCY
    Call copySheet(SHEET_WITH_VACANCY, SHEET_WITHOUT_VACANCY)
    
    Sheets(SHEET_WITHOUT_VACANCY).Select
    
    Range("A1").Select
    Dim endRow As Long
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' находим первую пустую строку
    Dim nullNumber As Long
    nullNumber = 2
    Do While Cells(nullNumber, 2) <> ""
        nullNumber = nullNumber + 1
    Loop
    
    ' оставляем в фильтре только строчки без табельного
    Rows("1:1").Select
    Selection.AutoFilter Field:=2, Criteria1:="="
    
    ' и удаляем, начиная с первой строки
    Range("A" + CStr(nullNumber) + ":H" + CStr(nullNumber)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Rows("1:1").Select
    Selection.AutoFilter
    
End Sub

'
Private Sub copySheet(fromSheet As String, toSheet As String)

    Sheets(fromSheet).Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.copy
    
    Sheets(toSheet).Select
    Range("A1").Select
    ActiveSheet.Paste

End Sub

'
Private Sub alignColumns(sheet As String)

    Sheets(sheet).Select
    Columns("A:J").EntireColumn.AutoFit

End Sub

Private Sub addRate(sheet As String)

    Const RATE = 9
    Const POSITION = 5

    Sheets(sheet).Select
    Cells(1, RATE) = "Ставки"
    
    Range("A1").Select
    Dim endRow As Long
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For beginRow = 2 To endRow Step 1
        If InStr(Cells(beginRow, POSITION), "парт-") = 0 Then
            Cells(beginRow, RATE) = "1"
        Else
            Cells(beginRow, RATE) = "0.5"
        End If
    Next beginRow
    
    ' копируем формат ячейки, как у E2
    Range("E2").Select
    Selection.copy
    Range("I2:I" + CStr(endRow)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' копируем формат ячейки, как у E1
    Range("E1").Select
    Selection.copy
    Range("I1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

End Sub

Private Sub addSubdivision(sheet As String)

    Const RATE = 10

    Sheets(sheet).Select
    Cells(1, RATE) = "Подразделение"
    
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],справочник!C[-9]:C[-8],2,0)"
    
    ' копируем формат ячейки, как у E2
    Range("E2").Select
    Selection.copy
    Range("J2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' смотрим, сколько у нас строк
    Range("A1").Select
    Dim endRow As Long
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' протягиваем формулу до последней строки
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J" + CStr(endRow))
    Range("J2:J" + CStr(endRow)).Select
    
    ' вставляем, как значения
    Columns("J:J").Select
    Selection.copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ' копируем формат ячейки, как у E1
    Range("E1").Select
    Selection.copy
    Range("J1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub