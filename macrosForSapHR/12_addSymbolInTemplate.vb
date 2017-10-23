' Признак "A" - сокращенный день. Признак "F" - праздничный день.
' Если день отмечен, как выходной (FREE), то признак сокращенного дня ("A") для него не проставляется.
'
Private Sub addSymbolInTemplate(dayNumber As Integer, index As Integer)
Worksheets("myCode").Select

Dim row, col As Integer
col = 3 + dayNumber * 5 - index ' 3 заголовка + (количество дней * 5 граф для каждого дня) - количество смещений

    For row = 2 To ALL_GRAPHS_AND_ONE_HEADER Step 1
        If index = 1 Then
            If Cells(row, col - 3) <> "FREE" Then
                Cells(row, col) = "A"
            End If
        Else
            If index = 2 Then
                Cells(row, col) = "1"
            End If
            If row >= 2 And row <= ALL_STADARD_GRAPHS_AND_ONE_HEADER And Cells(row, col + index - 4) <> "FREE" Then
                Cells(row, col + index - 1) = "F"
            End If
        End If
    Next row
    
End Sub