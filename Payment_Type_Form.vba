Private Sub UserForm_Initialize()

    With Me.Payment_Type_Combobox
    
            .ControlTipText = "Выберите значение из списка"
            .List = Array("Платёж успешен", "Возврат успешен", "Переводы", "Выписки")
    End With

End Sub

Private Sub CommandButton1_Click()

    Unload Me
    
    Select Case Payment_Type_Combobox.Text
    
        Case Is = "Платёж успешен"
            Payment_Success.Show
            
        Case Is = "Возврат успешен"
            Payment_Refund.Show
            
        Case Is = "Переводы"
            Payment_Transfer.Show
            
        Case Is = "Выписки"
            Official_DocForm.Show
        
        Case Is = "Отмена/Ошибка авторизации"
            Payment_Auth.Show
        
    End Select
    
End Sub