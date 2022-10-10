
Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()
    With Me.ComboBox1
        .ControlTipText = "Выберите значение из списка"
        .List = Array("Перевод p2p успешен", "Перевод c2c успешен", "Перевод на БК успешен")
    End With
    Money_Value_Kop.Enabled = False
    Money_Value_Kop.Text = "00"
    NKO_Money_Value_Kop.Enabled = False
    NKO_Money_Value_Kop.Text = "00"
End Sub



Private Sub ComboBox1_Change()

Dim Default_Value As String
Default_Value = "Заполнение не требуется"
        
    Select Case ComboBox1.Text
    
        Case Is = "Перевод на БК успешен"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            Money_Value.Enabled = True
            Money_Value.Text = ""
            NKO_Money_Value.Enabled = True
            NKO_Money_Value.Text = ""
            NKO_Money_Value_Kop.Enabled = False
            NKO_Money_Value_Kop.Text = "00"
            RRN.Enabled = True
            RRN.Text = ""
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = False
            NKO_Comission.Text = Default_Value
            
        Case Is = "Перевод p2p успешен"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            Payment_ID.Enabled = False
            Payment_ID.Text = Default_Value
            Money_Value.Enabled = True
            Money_Value.Text = ""
            NKO_Money_Value.Enabled = False
            NKO_Money_Value.Text = Default_Value
            NKO_Money_Value_Kop.Enabled = False
            NKO_Money_Value_Kop.Text = Default_Value
            RRN.Enabled = False
            RRN.Text = Default_Value
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = True
            NKO_Comission.Text = ""
            
        Case Is = "Перевод c2c успешен"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            Payment_ID.Enabled = False
            Payment_ID.Text = Default_Value
            Money_Value.Enabled = True
            Money_Value.Text = ""
            NKO_Money_Value.Enabled = False
            NKO_Money_Value.Text = Default_Value
            NKO_Money_Value_Kop.Enabled = False
            NKO_Money_Value_Kop.Text = Default_Value
            RRN.Enabled = True
            RRN.Text = ""
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = False
            NKO_Comission.Text = Default_Value
            
            
    End Select
End Sub

Private Sub CheckBox1_Change()
    Select Case CheckBox1.Value
        Case Is = True
            Money_Value_Kop.Text = " "
            Money_Value_Kop.Enabled = True
        
        Case Is = False
            Money_Value_Kop.Text = "00"
            Money_Value_Kop.Enabled = False
    End Select
End Sub

Private Sub CheckBox2_Change()
    If ComboBox1.Text = "Перевод на БК успешен" Then
        Select Case CheckBox2.Value
            Case Is = True
                
                    NKO_Money_Value_Kop.Text = " "
                    NKO_Money_Value_Kop.Enabled = True
            
            Case Is = False
                NKO_Money_Value_Kop.Text = "00"
                NKO_Money_Value_Kop.Enabled = False
        End Select
    Else:   NKO_Money_Value_Kop.Enabled = False And NKO_Money_Value_Kop.Text = Default_Value
            
    End If
End Sub

Private Sub CommandButton1_Click()
    If EmptyCheck() = True Then
    
            Call Add_Data(2, Ticket_Number.Text)
            Call Add_Data(3, ComboBox1.Text)
            Call Add_DV_Number(4, Ticket_Number.Text)
            Call Add_Data(5, Card_Number.Text)
            Call Add_Data(6, Date_Value.Text)
            Call Add_Data(9, KA_Value.Text)
            Call Add_Data(11, Payment_ID.Text)
            Call Add_Data(12, Money_Value.Text)
            Call Add_Data(13, Money_Value_Kop.Text)
            Call Add_Data(15, RRN.Text)
            Call Add_Data(17, NKO_Money_Value.Text)
            Call Add_Data(18, NKO_Money_Value_Kop.Text)
            Call Add_Data(19, NKO_Comission.Text)
            
            If PDF_Check_Box.Value = True Then
                Worksheets("Data").Cells(2, 7).Value = "1"
                Else: Worksheets("Data").Cells(2, 7).Value = "0"
            End If
        
            MsgBox ("Успех! Далее нажми кнопку 'Сформировать подтверждение'")
            
            Unload Payment_Transfer
            
    Else: MsgBox ("Необходимо заполнить след. поля: Номер тикета, Вид подтверждения, Номер карты, Дата")
    End If
End Sub


