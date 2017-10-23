' Старт макроса
'
Sub myWorkingProject()

ThisWorkbook.Save
Worksheets("directory").Visible = 2
Worksheets("directory_d").Visible = 2
Worksheets("directory_n").Visible = 2
Worksheets("architectureMacro").Visible = 2
Worksheets("selection").Visible = 2

Dim endRow As Long
endRow = Cells(Rows.Count, 1).End(xlUp).row
ALL_GRAPHS = endRow - HEADER_ROW - HEADER_ROW

ALL_GRAPHS_AND_ONE_HEADER = ALL_GRAPHS + HEADER_ROW
ALL_GRAPHS_AND_TWO_HEADER = ALL_GRAPHS + HEADER_ROW + HEADER_ROW
ALL_DAYS_AND_ONE_HEADER = ALL_DAYS + HEADER_ROW
ALL_DAYS_AND_TWO_HEADER = ALL_DAYS + HEADER_ROW + HEADER_ROW
ALL_STADARD_GRAPHS_AND_ONE_HEADER = ALL_STADARD_GRAPHS + HEADER_ROW

If checkColors() Then
Else: Exit Sub
End If

Call copyUserInput
Call selectionOfWorkSchedules
Call deleteOldData
Call fillPatternForSAP
Call addShortDayAndHoliday

Worksheets("forUser").Select
ThisWorkbook.Save
MsgBox "Все готово! Жми OK!", vbExclamation, "version: 006.000.000"

End Sub