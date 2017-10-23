' Разрешаем действия пользователя
'
Private Sub allowUserActions()

Application.ScreenUpdating = True
Application.Interactive = True
Application.EnableCancelKey = xlInterrupt
    
End Sub