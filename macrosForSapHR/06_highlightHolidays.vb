' Выделяем праздничные дни (с признаком "F") оранжевым для отладки
'
Private Sub highlightHolidays()
Worksheets("selection").Select

Dim row, col As Integer
For row = 2 To ALL_GRAPHS_AND_ONE_HEADER Step 1
    For col = 3 To ALL_DAYS_AND_TWO_HEADER Step 1
        If Cells(row, col) = "F" Then
            Cells(row, col).Interior.color = ORANGE
        End If
    Next col
Next row

End Sub