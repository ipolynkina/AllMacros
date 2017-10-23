' ************************* generate report on region *************************
'
Public Const RR_BEGINNING_OF_DATA = 3
Public Const RR_NAME_LIST = "regions"

Private Const RR_ID_REGION = 1
Private Const RR_COL_FIRST_MONTH = 2
Private Const RR_COL_LAST_MONTH = 37

Private Const RR_BLOCK_BEGIN = "A"
Private Const RR_BLOCK_END = "AK"
Private Const BLUE = 14994616

Private Const FIRST_MONTH = 1
Private Const LAST_MONTH = 12
'
Public Sub generateReportOnRegions(isOpenFromOutside As Boolean)

    Call copyOriginalData(True)
    Call setRegions(Main.NAME_NEW_LIST)
    Call setDivisions(Main.NAME_NEW_LIST)
    Call consolidateDataOnRegions(Main.NAME_NEW_LIST, RR_NAME_LIST)
    Call setFilterOnColumn(RR_NAME_LIST, RR_BEGINNING_OF_DATA - 1, RR_BLOCK_BEGIN)
    Call setResults(RR_NAME_LIST)
    Range("A1").Select

End Sub
'
Private Sub consolidateDataOnRegions(sheetNameForAnalyze As String, sheetNameForWrite As String)

    Sheets(sheetNameForAnalyze).Select
    Dim row As Integer, rowForWrite As Integer, i As Integer
    Dim amountSellers(FIRST_MONTH To LAST_MONTH) As Integer, amountPartTimers(FIRST_MONTH To LAST_MONTH) As Integer
    Dim region As String
    
    row = Main.BEGINNING_OF_DATA
    rowForWrite = RR_BEGINNING_OF_DATA
    
    Do While Cells(row, Main.ID_REGION) <> ""
        For i = FIRST_MONTH To LAST_MONTH Step 1
            amountSellers(i) = 0
            amountPartTimers(i) = 0
        Next i
        
        region = Cells(row, Main.ID_REGION)
        Do While Cells(row, Main.ID_REGION) <> ""
            If Cells(row, Main.ID_REGION) = region Then
                If Cells(row, Main.ID_POSITION) = Main.SIGN_SELLER Then
                    amountSellers(Cells(row, Main.ID_MONTH)) = amountSellers(Cells(row, Main.ID_MONTH)) + 1
                Else:
                    amountPartTimers(Cells(row, Main.ID_MONTH)) = amountPartTimers(Cells(row, Main.ID_MONTH)) + 1
                End If
                deleteRow (row)
                row = row - 1
            End If
            row = row + 1
        Loop

        Call writeDataOnRegion(sheetNameForAnalyze, sheetNameForWrite, rowForWrite, region, amountSellers(), amountPartTimers())
        rowForWrite = rowForWrite + 1
        row = Main.BEGINNING_OF_DATA
    Loop

End Sub
'
Private Sub writeDataOnRegion(sheetNameForAnalyze As String, sheetNameForWrite As String, rowForWrite As Integer, _
                              region As String, amountSellers() As Integer, amountPartTimers() As Integer)
    
    Sheets(sheetNameForWrite).Select
    Const SIZE_STEP = 1
    Dim columnForWrite As Integer, i As Integer
    columnForWrite = RR_COL_FIRST_MONTH
    
    Cells(rowForWrite, RR_ID_REGION) = region
    For i = FIRST_MONTH To LAST_MONTH Step 1
        Cells(rowForWrite, columnForWrite) = amountSellers(i)
        columnForWrite = columnForWrite + SIZE_STEP
        
        Cells(rowForWrite, columnForWrite) = amountPartTimers(i)
        columnForWrite = columnForWrite + SIZE_STEP
        
        Cells(rowForWrite, columnForWrite) = amountSellers(i) + amountPartTimers(i)
        columnForWrite = columnForWrite + SIZE_STEP
    Next i
    
    Sheets(sheetNameForAnalyze).Select
    
End Sub
'
Private Sub setResults(sheetName As String)

    Sheets(sheetName).Select
    Dim row As Integer, nextRow As Integer, sum As Integer, indexMonth As Integer
    nextRow = Cells(Rows.Count, 1).End(xlUp).row + 1
    sum = 0
    
    Cells(nextRow, RR_ID_REGION) = "Общий итог"
    For indexMonth = RR_COL_FIRST_MONTH To RR_COL_LAST_MONTH Step 1
        For row = RR_BEGINNING_OF_DATA To nextRow - 1 Step 1
            sum = sum + Cells(row, indexMonth)
        Next row
        Cells(nextRow, indexMonth) = sum
        sum = 0
    Next indexMonth
    
    Call setColor(nextRow, BLUE, RR_BLOCK_BEGIN, RR_BLOCK_END)

End Sub