' Значение сокращенного дня (они выделены фиолетовым) увеличиваем на 1.
' Дни, выделенные полужирным шрифтом, обозначают ночные смены (обычный шрифт - дневные смены).
'
Private Sub selectionOfWorkSchedules()

Call searchSchedules
Call highlightHolidays
    
Dim row As Integer
Dim col As Integer
Dim currColor As Variant
Const BIAS As Integer = 44

For row = 2 To ALL_GRAPHS_AND_ONE_HEADER Step 1
    For col = 3 To ALL_DAYS_AND_TWO_HEADER Step 1
        currColor = Cells(row, col).Interior.color
        If currColor = PURPLE Then
            Cells(row, col) = Cells(row, col) + 1
            If Selection.Font.Bold = True Then ' полужирный
                Cells(row, col).Interior.color = BLUE
            Else: Cells(row, col).Interior.color = YELLOW
            End If
            currColor = Cells(row, col).Interior.color
        End If
        Cells(row, col + BIAS).Select
        Call highlightWeekdays(currColor, row, col)
                If Cells(row, col + BIAS) = 0 Then
                    Cells(row, col + BIAS) = "FREE"
                    Cells(row, col + BIAS).Interior.color = NO_COLOR
                End If
    Next col
Next row

End Sub