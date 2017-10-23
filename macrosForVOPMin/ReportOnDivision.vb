' ************************* generate report on division *************************
'
Public Const RD_BEGINNING_OF_DATA = 2
Public Const RD_NAME_LIST = "divisions"

Private Const RD_ID_REGION = 1
Private Const RD_ID_DIVISION = 2
Private Const RD_ID_FIRST_MONTH = 3
Private Const RD_ID_LAST_MONTH = 14

Private Const RD_BLOCK_BEGIN = "A"
Private Const RD_BLOCK_END = "N"
Private Const BLUE = 14994616

Private Const FIRST_MONTH = 1
Private Const LAST_MONTH = 12
'
Public Sub generateReportOnDivisions(isOpenFromOutside As Boolean)

    Call copyOriginalData(True)
    Call setRegions(Main.NAME_NEW_LIST)
    Call setDivisions(Main.NAME_NEW_LIST)
    Call consolidateDataOnDivisions(Main.NAME_NEW_LIST, RD_NAME_LIST)
    Call setFilterOnColumn(RD_NAME_LIST, RD_BEGINNING_OF_DATA - 1, RD_BLOCK_BEGIN)
    Call setResults(RD_NAME_LIST)
    Range("A1").Select

End Sub
'
Private Sub consolidateDataOnDivisions(sheetNameForAnalyze As String, sheetNameForWrite As String)

    Sheets(sheetNameForAnalyze).Select
    Dim row As Integer, rowForWrite As Integer, i As Integer
    Dim amountAllSellers(FIRST_MONTH To LAST_MONTH) As Integer
    Dim region As String, division As String
    
    row = RD_BEGINNING_OF_DATA
    rowForWrite = RD_BEGINNING_OF_DATA
    
    Do While Cells(row, Main.ID_DIVISION) <> ""
        For i = FIRST_MONTH To LAST_MONTH Step 1
            amountAllSellers(i) = 0
        Next i

        region = Cells(row, Main.ID_REGION)
        division = Cells(row, Main.ID_DIVISION)
        
        Do While Cells(row, Main.ID_DIVISION) <> ""
            If Cells(row, Main.ID_DIVISION) = division Then
                amountAllSellers(Cells(row, Main.ID_MONTH)) = amountAllSellers(Cells(row, Main.ID_MONTH)) + 1
                deleteRow (row)
                row = row - 1
            End If
            row = row + 1
        Loop

        Call writeDataOnDivision(sheetNameForAnalyze, sheetNameForWrite, rowForWrite, region, division, amountAllSellers())
        rowForWrite = rowForWrite + 1
        row = RD_BEGINNING_OF_DATA
    Loop

End Sub
'
Private Sub writeDataOnDivision(sheetNameForAnalyze As String, sheetNameForWrite As String, rowForWrite As Integer, _
                                region As String, division As String, amountAllSellers() As Integer)

    Sheets(sheetNameForWrite).Select
    Dim columnForWrite As Integer, i As Integer
    columnForWrite = RD_ID_FIRST_MONTH
    
    Cells(rowForWrite, RD_ID_REGION) = region
    Cells(rowForWrite, RD_ID_DIVISION) = division
    
    For i = FIRST_MONTH To LAST_MONTH Step 1
        Cells(rowForWrite, columnForWrite) = amountAllSellers(i)
        columnForWrite = columnForWrite + 1
    Next i
    
    Sheets(sheetNameForAnalyze).Select

End Sub
'
Private Sub setResults(sheetName As String)
    
    Sheets(sheetName).Select
    Dim endRow As Integer
    endRow = Cells(Rows.Count, 1).End(xlUp).row
    
    Call addResults(sheetName, endRow + 1, "Общий", RD_BEGINNING_OF_DATA, endRow)
    Call addResultsByRegions(sheetName)

End Sub
'
Private Sub addResults(sheetName As String, rowForWrite As Integer, text As String, startIndex As Integer, endIndex As Integer)

    Sheets(sheetName).Select
    Dim row As Integer, sum As Integer, indexMonth As Integer
    sum = 0
    
    Cells(rowForWrite, RD_ID_REGION) = text + " итог"
    For indexMonth = RD_ID_FIRST_MONTH To RD_ID_LAST_MONTH Step 1
        For row = startIndex To endIndex Step 1
            sum = sum + Cells(row, indexMonth)
        Next row
        Cells(rowForWrite, indexMonth) = sum
        sum = 0
    Next indexMonth
    
    Call setColor(rowForWrite, BLUE, RD_BLOCK_BEGIN, RD_BLOCK_END)
    
End Sub
'
Private Sub addResultsByRegions(sheetName As String)

    Sheets(sheetName).Select
    Dim row As Integer, endRow As Integer, startIndex As Integer
    Dim region As String
    
    row = RD_BEGINNING_OF_DATA
    endRow = Cells(Rows.Count, 1).End(xlUp).row
    startIndex = RD_BEGINNING_OF_DATA
    region = Cells(row, RD_ID_REGION)
    
    Do While region <> "Общий итог"
        Do While Cells(row, RD_ID_REGION) = region
            row = row + 1
        Loop
        
        Rows(CStr(row) + ":" + CStr(row)).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Call addResults(sheetName, row, region, startIndex, row - 1)
        row = row + 1
        startIndex = row
        region = Cells(row, RD_ID_REGION)
    Loop

End Sub