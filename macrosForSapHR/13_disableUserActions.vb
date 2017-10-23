' Отключаем дейстивия пользователя
'
Private Sub disableUserActions()

Application.ScreenUpdating = False
Application.Interactive = False
Application.EnableCancelKey = xlDisabled
    
End Sub