Public Const VERSION = "version 2.0"
Public Const SHEET_KM = "обр_без_вак_МАГ"
Public Const SHEET_IM = "обр_без_вак_ИМ"

'
Sub plus_km_and_im()

    ' если есть #Н/Д - дальше не обрабатываем
    If vprOK() Then
    Else: Exit Sub
    End If

    Call deleteSheets
    Call createSheetKM
    Call createSheetIM
    
    Call alignColumns(SHEET_KM)
    Sheets(SHEET_KM).Select
    Range("A1").Select
    
    Call alignColumns(SHEET_IM)
    Sheets(SHEET_IM).Select
    Range("A1").Select
    
    ThisWorkbook.Save
    MsgBox "Обработка успешно завершена", vbExclamation, VERSION
    
End Sub

'
Private Function vprOK()

    Sheets(SHEET_WITHOUT_VACANCY).Select
    
    Range("A1").Select
    Dim endRow As Long
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$" + CStr(endRow)).AutoFilter Field:=10, Criteria1:="#Н/Д"
    
    Range("A1").Select
    Dim amountStr As Long
    amountStr = Cells(Rows.Count, 1).End(xlUp).Row
    
    Rows("1:1").Select
    Selection.AutoFilter
    
    If amountStr > 1 Then
        vprOK = False
        MsgBox "Заполните вручную Подразделение" & vbCrLf & "И запустите макрос plus_km_and_im еще раз", vbExclamation, VERSION
    Else
        vprOK = True
    End If

End Function

'
Private Sub deleteSheets()

    Application.DisplayAlerts = False
    
    Dim i As Long
    For i = Sheets.Count To 1 Step - 1
        If Sheets(i).Name = SHEET_KM _
        Or Sheets(i).Name = SHEET_IM Then
            Sheets(i).Delete
        End If
    Next
    
    Application.DisplayAlerts = True

End Sub

'
Private Sub createSheetKM()

    Sheets.Add.Name = SHEET_KM
    Sheets(SHEET_WITHOUT_VACANCY).Select
    
    Range("A1").Select
    Dim endRow As Long
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$" + CStr(endRow)).AutoFilter Field:=10, Criteria1:="Магазин"
    Range("A1:J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.copy
    
    Sheets(SHEET_KM).Select
    ActiveSheet.Paste
    
    Sheets(SHEET_WITHOUT_VACANCY).Select
    Rows("1:1").Select
    Selection.AutoFilter

End Sub

'
Private Sub createSheetIM()

    Sheets.Add.Name = SHEET_IM
    Sheets(SHEET_WITHOUT_VACANCY).Select
    
    Range("A1").Select
    Dim endRow As Long
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$" + CStr(endRow)).AutoFilter Field:=10, Criteria1:="ИМ"
    Range("A1:J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.copy
    
    Sheets(SHEET_IM).Select
    ActiveSheet.Paste
    
    Sheets(SHEET_WITHOUT_VACANCY).Select
    Rows("1:1").Select
    Selection.AutoFilter

End Sub

'
Private Sub alignColumns(sheet As String)

    Sheets(sheet).Select
    Columns("A:J").EntireColumn.AutoFit

End Sub