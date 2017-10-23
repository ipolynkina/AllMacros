' Добавляем в шаблон признаки выходных и сокращенных дней
'
Private Sub addShortDayAndHoliday()
Worksheets("myCode").Select

Dim index As Integer
Dim dayNumber As Integer
Dim messageText(1 To 5) As String
messageText(1) = "СОКРАЩЕННЫЙ ДЕНЬ (введите дату или 0): "
messageText(2) = "ПРАЗДНИЧНЫЙ ДЕНЬ (введите дату или 0): "
messageText(3) = "ПЕРЕНЕСЕННЫЙ ПРАЗДНИЧНЫЙ ДЕНЬ (введите дату или 0): "
messageText(4) = "Ну где ж Вы даты такие видели?"
messageText(5) = "P.S. Для выхода нажмите 0"

Const MAX_AMOUNT_STEPS As Integer = 3
For index = 1 To MAX_AMOUNT_STEPS Step 1
    dayNumber = InputBox(messageText(index), messageText(5))
    Do While dayNumber <> 0
        If dayNumber > 0 And dayNumber <= ALL_DAYS Then
            Call addSymbolInTemplate(dayNumber, index)
        Else
            MsgBox messageText(4), vbQuestion, "Error"
        End If
        dayNumber = InputBox(messageText(index), messageText(5))
    Loop
Next index

End Sub