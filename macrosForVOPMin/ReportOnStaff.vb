' ************************* generate report on staff *************************
'
Public Const RS_BEGINNING_OF_DATA = 2
Public Const RS_NAME_LIST = "employees"

Private Const RS_ID_SHOP = 1
Private Const RS_ID_REGION = 2
Private Const RS_ID_DIVISION = 3
Private Const RS_ID_PERS_NUM = 4
Private Const RS_ID_FILENAME = 5
Private Const RS_ID_POSITION = 6
Private Const RS_ID_AMOUNT_OCCURRENCES = 19
Private Const RS_BLOCK_BEGIN = "A"
Private Const RS_BLOCK_END = "S"

Private Const SIGN_BLOCK_FINAL_ENTRY = "S"

Private Const COLOR_FOR_STAFF_WITH_1_VIOLATIONS = 14806254
Private Const COLOR_FOR_STAFF_WITH_2_VIOLATIONS = 11851260
Private Const COLOR_FOR_STAFF_WITH_3_VIOLATIONS = 12040422
Private Const COLOR_FOR_STAFF_WITH_4_VIOLATIONS = 49407
Private Const COLOR_FOR_STAFF_WITH_5_VIOLATIONS = 13311
Private Const COLOR_FOR_STAFF_WITH_6_VIOLATIONS = 255
Private Const COLOR_FOR_STAFF_WITH_7_VIOLATIONS = 192
'
Public Sub generateReportOnStaff(isOpenFromOutside As Boolean)
    
    Call copyOriginalData(True)
    Call setRegions(Main.NAME_NEW_LIST)
    Call setDivisions(Main.NAME_NEW_LIST)
    Call consolidateDataOnEmployees(Main.NAME_NEW_LIST, RS_NAME_LIST)
    Call countAmountOfOccurrences(RS_NAME_LIST)
    Call setFilterOnColumn(RS_NAME_LIST, RS_BEGINNING_OF_DATA - 1, SIGN_BLOCK_FINAL_ENTRY)
    Range("A1").Select

End Sub
'
Private Sub consolidateDataOnEmployees(sheetNameForAnalyze As String, sheetNameForWrite As String)
    
    Sheets(sheetNameForAnalyze).Select
    Dim persNum As Long, vop As Long
    Dim row As Integer, rowForWrite As Integer, indexMonth As Integer
    Dim shop As String, filename As String, position As String, region As String, division As String
    
    row = BEGINNING_OF_DATA
    rowForWrite = BEGINNING_OF_DATA
    
    Do While Cells(row, Main.ID_PERS_NUM) <> ""
        shop = Cells(row, Main.ID_SHOP)
        region = Cells(row, Main.ID_REGION)
        division = Cells(row, Main.ID_DIVISION)
        persNum = Cells(row, Main.ID_PERS_NUM)
        filename = Cells(row, Main.ID_FILENAME)
        position = Cells(row, Main.ID_POSITION)
        vop = Cells(row, Main.ID_VOP)
        indexMonth = Cells(row, Main.ID_MONTH)
        deleteRow (row)
        Call addEmployee(sheetNameForAnalyze, sheetNameForWrite, rowForWrite, _
                         shop, region, division, persNum, filename, position, vop, indexMonth)
        Do While Cells(row, Main.ID_PERS_NUM) <> ""
            If Cells(row, Main.ID_PERS_NUM) = persNum And Cells(row, Main.ID_SHOP) = shop Then
                filename = Cells(row, Main.ID_FILENAME)
                position = Cells(row, Main.ID_POSITION)
                indexMonth = Cells(row, Main.ID_MONTH)
                vop = Cells(row, Main.ID_VOP)
                deleteRow (row)
                row = row - 1
                Call updateEmployeeData(sheetNameForAnalyze, sheetNameForWrite, rowForWrite, _
                                        filename, position, vop, indexMonth)
            End If
            row = row + 1
        Loop
        
        row = BEGINNING_OF_DATA
        rowForWrite = rowForWrite + 1
    Loop

End Sub
'
Private Sub addEmployee(sheetNameForAnalyze As String, sheetNameForWrite As String, row As Integer, _
                        shop As String, region As String, division As String, persNum As Long, _
                        filename As String, position As String, vop As Long, indexMonth As Integer)

    Sheets(sheetNameForWrite).Select
    Cells(row, RS_ID_SHOP) = shop
    Cells(row, RS_ID_REGION) = region
    Cells(row, RS_ID_DIVISION) = division
    Cells(row, RS_ID_PERS_NUM) = persNum
    Cells(row, RS_ID_FILENAME) = filename
    Cells(row, RS_ID_POSITION) = position
    Cells(row, RS_ID_POSITION + indexMonth) = vop
    Sheets(sheetNameForAnalyze).Select

End Sub
'
Private Sub updateEmployeeData(sheetNameForAnalyze As String, sheetNameForWrite As String, row As Integer, _
                               filename As String, position As String, vop As Long, indexMonth As Integer)

    Sheets(sheetNameForWrite).Select
    Cells(row, RS_ID_FILENAME) = filename
    Cells(row, RS_ID_POSITION) = position
    Cells(row, RS_ID_POSITION + indexMonth) = vop
    Sheets(sheetNameForAnalyze).Select

End Sub
'
Private Sub countAmountOfOccurrences(sheetName As String)

    Sheets(sheetName).Select
    Dim row As Integer, endRow As Integer
    Dim codeColor As Long
    endRow = Cells(Rows.Count, 1).End(xlUp).row
    For row = BEGINNING_OF_DATA To endRow Step 1
        Cells(row, RS_ID_AMOUNT_OCCURRENCES).Value = "=COUNT(RC[-12]:RC[-1])"
        codeColor = getCodeColor(Cells(row, RS_ID_AMOUNT_OCCURRENCES))
        Call setColor(row, codeColor, RS_BLOCK_BEGIN, RS_BLOCK_END)
    Next row
     
    Columns("S:S").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Sub
'
Private Function getCodeColor(amountViolations As Integer) As Long
    If amountViolations = 1 Then
        getCodeColor = COLOR_FOR_STAFF_WITH_1_VIOLATIONS
    ElseIf amountViolations = 2 Then
        getCodeColor = COLOR_FOR_STAFF_WITH_2_VIOLATIONS
    ElseIf amountViolations = 3 Then
        getCodeColor = COLOR_FOR_STAFF_WITH_3_VIOLATIONS
    ElseIf amountViolations = 4 Then
        getCodeColor = COLOR_FOR_STAFF_WITH_4_VIOLATIONS
    ElseIf amountViolations = 5 Then
        getCodeColor = COLOR_FOR_STAFF_WITH_5_VIOLATIONS
    ElseIf amountViolations = 6 Then
        getCodeColor = COLOR_FOR_STAFF_WITH_6_VIOLATIONS
    Else
        getCodeColor = COLOR_FOR_STAFF_WITH_7_VIOLATIONS
    End If
End Function