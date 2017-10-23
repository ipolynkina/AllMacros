' macro created for data integration between BW and SAP HR
' author: Polynkina Irina
' contact: irina.polynkina.dev@yandex.ru
' version: 1.1.0
' release: 13.12.2016
'
Const BEGINNING_OF_DATA = 2

Const SIGN_BLOCK_BW = "A"
Const SIGN_SHOPS_BW = "A"
Const SIGN_PERS_NUM_BW = "C"
Const ID_SHOP_BW = 1
Const ID_PERS_NUM_BW = 3
Const ID_FILENAME_BW = 5
Const ID_VOLUME_OF_SALES_BW = 6
Const ID_EXCLUDED_SALES_BW = 7

Const SIGN_BLOCK_SAP = "N"
Const SIGN_SHOPS_SAP = "Q"
Const SIGN_PERS_NUM_SAP = "N"
Const ID_PERS_NUM_SAP = 14
Const ID_SHOP_SAP = 17
Const ID_BEGIN_DATE_SAP = 18
Const ID_END_DATE_SAP = 19
'
' start macro
'
Sub beginIntegration()

    ThisWorkbook.Save
    Call disableUserActions
    Worksheets("employees").Select
    
    Call sortData(SIGN_BLOCK_BW, SIGN_SHOPS_BW)
    Call sortData(SIGN_BLOCK_BW, SIGN_PERS_NUM_BW)
    
    Call sortData(SIGN_BLOCK_SAP, SIGN_SHOPS_SAP)
    Call sortData(SIGN_BLOCK_SAP, SIGN_PERS_NUM_SAP)
    
    Call combineDataInSAP
    Call combineDataInBW
    Call dataIntegration
    
    ThisWorkbook.Save
    Call allowUserActions
    
    MsgBox "Все готово! Жми OK!", vbExclamation, "version: 1.1.0"
    
End Sub
'
Private Sub sortData(signOfBlock As String, signFieldForSorting As String)
    
    Range(signOfBlock + "1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("employees").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("employees").AutoFilter.Sort.SortFields.Add Key:= _
        Range(signFieldForSorting + "1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("employees").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
End Sub
'
Private Sub combineDataInSAP()

    Dim row As Integer, rowForDelete As Integer
    row = BEGINNING_OF_DATA
     
    Do While Cells(row, ID_PERS_NUM_SAP) <> ""
        If Cells(row, ID_PERS_NUM_SAP) = Cells(row + 1, ID_PERS_NUM_SAP) And _
           Cells(row, ID_SHOP_SAP) = Cells(row + 1, ID_SHOP_SAP) And _
           Cells(row, ID_END_DATE_SAP) + 1 = Cells(row + 1, ID_BEGIN_DATE_SAP) Then
                Cells(row, ID_BEGIN_DATE_SAP) = getMinDate(Cells(row, ID_BEGIN_DATE_SAP), Cells(row + 1, ID_BEGIN_DATE_SAP))
                Cells(row, ID_END_DATE_SAP) = getMaxDate(Cells(row, ID_END_DATE_SAP), Cells(row + 1, ID_END_DATE_SAP))
                rowForDelete = row + 1
                Range(SIGN_BLOCK_SAP + CStr(rowForDelete)).Select
                Range(Selection, Selection.End(xlToRight)).Select
                Selection.Delete Shift:=xlUp
                row = row - 1
        End If
        row = row + 1
    Loop

End Sub
'
Private Sub combineDataInBW()

    Dim row As Integer, rowForDelete As Integer
    row = BEGINNING_OF_DATA
   
    Do While Cells(row, ID_PERS_NUM_BW) <> ""
        If Cells(row, ID_PERS_NUM_BW) = Cells(row + 1, ID_PERS_NUM_BW) And _
           Cells(row, ID_SHOP_BW) = Cells(row + 1, ID_SHOP_BW) And _
           Cells(row, ID_FILENAME_BW) = Cells(row + 1, ID_FILENAME_BW) Then
                Cells(row, ID_VOLUME_OF_SALES_BW) = Cells(row, ID_VOLUME_OF_SALES_BW) + Cells(row + 1, ID_VOLUME_OF_SALES_BW)
                Cells(row, ID_EXCLUDED_SALES_BW) = Cells(row, ID_EXCLUDED_SALES_BW) + Cells(row + 1, ID_EXCLUDED_SALES_BW)
                rowForDelete = row + 1
                Range(SIGN_BLOCK_BW + CStr(rowForDelete)).Select
                Range(Selection, Selection.End(xlToRight)).Select
                Selection.Delete Shift:=xlUp
                row = row - 1
        End If
        row = row + 1
    Loop

End Sub
'
Private Function getMinDate(date1 As Date, date2 As Date)

    If date1 < date2 Then
        getMinDate = date1
    Else
        getMinDate = date2
    End If

End Function
'
Private Function getMaxDate(date1 As Date, date2 As Date) As Date

    If date1 > date2 Then
        getMaxDate = date1
    Else
        getMaxDate = date2
    End If

End Function
'
Private Sub dataIntegration()

    Dim row As Integer, rowForDelete As Integer
    row = BEGINNING_OF_DATA
    
    Do While Cells(row, ID_PERS_NUM_SAP) <> ""
        If Cells(row, ID_PERS_NUM_BW) <> Cells(row, ID_PERS_NUM_SAP) Then
            If Cells(row, ID_PERS_NUM_SAP) = Cells(row - 1, ID_PERS_NUM_SAP) And _
               Cells(row, ID_SHOP_SAP) <> Cells(row - 1, ID_SHOP_SAP) Then
                If Cells(row - 1, ID_SHOP_BW) = Cells(row - 1, ID_SHOP_SAP) Then
                    rowForDelete = row
                Else:
                    rowForDelete = row - 1
                End If
                Range(SIGN_BLOCK_SAP + CStr(rowForDelete)).Select
                Range(Selection, Selection.End(xlToRight)).Select
                Selection.Delete Shift:=xlUp
            Else:
                If Cells(row, ID_PERS_NUM_BW) < Cells(row, ID_PERS_NUM_SAP) Then
                    Range(SIGN_BLOCK_SAP + CStr(row)).Select
                Else:
                    Range(SIGN_BLOCK_BW + CStr(row)).Select
                End If
                Range(Selection, Selection.End(xlToRight)).Select
                Selection.insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            End If
        End If
        row = row + 1
    Loop

End Sub
'
Private Sub disableUserActions()

    Application.ScreenUpdating = False
    Application.Interactive = False
    Application.EnableCancelKey = xlDisabled
    
End Sub
'
Private Sub allowUserActions()
 
    Application.ScreenUpdating = True
    Application.Interactive = True
    Application.EnableCancelKey = xlInterrupt

End Sub