' Макрос создан для заполнения шаблона графиков рабочего времени SAP HR.
' Пользователь заполняет первый лист значениями типа: "8", "7", ..., "4", а макрос
' подбирает однодневный график в соответствии с заданным значением и именем самого графика.
'
' Polynkina Irina
' irina.polynkina.dev@yandex.ru
'
' first release: 26.11.2014
' current release: 15.06.2016
' version: 0006.000.000
'
Const ALL_STADARD_GRAPHS As Integer = 7
Const HEADER_ROW As Integer = 1
Const ALL_DAYS As Integer = 31
Const BLUE As Variant = 15773696
Const ORANGE As Variant = 49407
Const YELLOW As Variant = 65535
Const PURPLE As Variant = 10498160
Const DARK_GREEN As Variant = 5296274
Const LIGHT_GREEN As Variant = 16777215
Const RED As Variant = 5287936
Const NO_COLOR As Variant = 255

Public ALL_GRAPHS As Integer
Public ALL_GRAPHS_AND_ONE_HEADER As Integer
Public ALL_GRAPHS_AND_TWO_HEADER As Integer
Public ALL_DAYS_AND_ONE_HEADER As Integer
Public ALL_DAYS_AND_TWO_HEADER As Integer
Public ALL_STADARD_GRAPHS_AND_ONE_HEADER As Integer