' Цвета обозначают смены: голубой - ночные, желтый - дневные, фиолетовый - сокращенные;
' красный - выходные дни; светло и темно-зеленый - для удобства пользователя.
' Другие цвета использовать нельзя (но можно оставлять ячейки без заливки).
'
Private Function checkColors()
Worksheets("forUser").Select

Dim row, col, errorCounter As Integer
errorCounter = 0

For row = 3 To ALL_GRAPHS_AND_TWO_HEADER Step 1
    For col = 3 To ALL_DAYS_AND_TWO_HEADER Step 1
        If Cells(row, col).Interior.color <> BLUE And _
        Cells(row, col).Interior.color <> YELLOW And _
        Cells(row, col).Interior.color <> PURPLE And _
        Cells(row, col).Interior.color <> DARK_GREEN And _
        Cells(row, col).Interior.color <> LIGHT_GREEN And _
        Cells(row, col).Interior.color <> RED And _
        Cells(row, col).Interior.color <> NO_COLOR Then
        MsgBox "Строка " & row, vbCritical, "Неустановленный цвет"
        errorCounter = errorCounter + 1
        End If
    Next col
Next row

If errorCounter = 0 Then
    checkColors = True
Else
    MsgBox "Ошибки необходимо исправить. Всего их: " & errorCounter, vbCritical, "Неудача!"
    checkColors = False
End If

End Function