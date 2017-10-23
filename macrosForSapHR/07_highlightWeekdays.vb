' Если значение пользователя совпадает с часами, подобранными в функции searchSchedules - подставляем соответствующий график.
' Названия допустимых графиков пользователь задает на листе "forUser".
'
Private Sub highlightWeekdays(color As Variant, row As Integer, col As Integer)
Worksheets("selection").Select

Const BIAS As Integer = 44
Select Case color
    Case BLUE
        ActiveCell.FormulaR1C1 = _
        "=IF(RC[-44]=RC78,RC34,IF(RC[-44]=RC79,RC36,IF(RC[-44]=RC80,RC38,IF(RC[-44]=RC81,RC40,IF(RC[-44]=RC82,RC42,IF(RC[-44]=RC83,RC44,0))))))"
        Cells(row, col + BIAS).Interior.color = BLUE
    Case YELLOW
        ActiveCell.FormulaR1C1 = _
        "=IF(RC[-44]=RC84,RC34,IF(RC[-44]=RC85,RC36,IF(RC[-44]=RC86,RC38,IF(RC[-44]=RC87,RC40,IF(RC[-44]=RC88,RC42,IF(RC[-44]=RC89,RC44,0))))))"
        Cells(row, col + BIAS).Interior.color = YELLOW
    Case ORANGE
        ActiveCell.FormulaR1C1 = "=RC34"
        Cells(row, col + BIAS).Interior.color = ORANGE
    Case LIGHT_GREEN, DARK_GREEN, RED, NO_COLOR
        ActiveCell.FormulaR1C1 = _
        "=(IF(RC35=RC[-44],RC34,IF(RC37=RC[-44],RC36,IF(RC39=RC[-44],RC38,IF(RC41=RC[-44],RC40,IF(RC43=RC[-44],RC42,IF(RC45=RC[-44],RC44,0)))))))"
        Cells(row, col + BIAS).Interior.color = LIGHT_GREEN
End Select

End Sub