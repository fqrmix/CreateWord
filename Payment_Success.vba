Private Sub UserForm_Initialize()
    With Me.ComboBox1
        .ControlTipText = "Выберите значение из списка"
        .List = Array("Платеж успешен 257", "Платеж успешен карта", "Платеж успешен инвойсинг", "Платеж успешен СБП", "Платеж успешен СБП наш КА", "Платеж успешен кошелек", "Платеж успешен ЮCard", "Платеж успешен с мобильного")
    End With
    Money_Value_Kop.Enabled = False
    Money_Value_Kop.Text = "00"
End Sub



Private Sub ComboBox1_Change()

Dim Default_Value As String
Default_Value = "Заполнение не требуется"
        
    Select Case ComboBox1.Text
    
        Case Is = "Платеж успешен 257"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            ID_Value.Enabled = True
            ID_Value.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            Order_Number.Enabled = False
            Order_Number.Text = Default_Value
            Auth_Code.Enabled = False
            Auth_Code.Text = Default_Value
            RRN.Enabled = False
            RRN.Text = Default_Value
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = True
            NKO_Comission.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
    
        Case Is = "Платеж успешен карта"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            ID_Value.Enabled = True
            ID_Value.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            Order_Number.Enabled = False
            Order_Number.Text = Default_Value
            Auth_Code.Enabled = False
            Auth_Code.Text = Default_Value
            RRN.Enabled = False
            RRN.Text = Default_Value
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = True
            NKO_Comission.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            
        Case Is = "Платеж успешен инвойсинг"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            ID_Value.Enabled = True
            ID_Value.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            Order_Number.Enabled = True
            Order_Number.Text = ""
            Auth_Code.Enabled = False
            Auth_Code.Text = Default_Value
            RRN.Enabled = False
            RRN.Text = Default_Value
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = False
            NKO_Comission.Text = Default_Value
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            
        Case Is = "Платеж успешен СБП"
            KA_Value.Enabled = False
            KA_Value.Text = Default_Value
            ID_Value.Enabled = False
            ID_Value.Text = Default_Value
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            Order_Number.Enabled = True
            Order_Number.Text = ""
            Auth_Code.Enabled = False
            Auth_Code.Text = Default_Value
            RRN.Enabled = False
            RRN.Text = Default_Value
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = False
            NKO_Comission.Text = Default_Value
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            
        Case Is = "Платеж успешен СБП наш КА"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            ID_Value.Enabled = False
            ID_Value.Text = Default_Value
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            Order_Number.Enabled = True
            Order_Number.Text = ""
            Auth_Code.Enabled = False
            Auth_Code.Text = Default_Value
            RRN.Enabled = False
            RRN.Text = Default_Value
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = False
            NKO_Comission.Text = Default_Value
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            
        Case Is = "Платеж успешен кошелек"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            ID_Value.Enabled = True
            ID_Value.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            Order_Number.Enabled = False
            Order_Number.Text = Default_Value
            Auth_Code.Enabled = False
            Auth_Code.Text = Default_Value
            RRN.Enabled = False
            RRN.Text = Default_Value
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = True
            NKO_Comission.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            
        Case Is = "Платеж успешен ЮCard"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            ID_Value.Enabled = False
            ID_Value.Text = Default_Value
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            Order_Number.Enabled = False
            Order_Number.Text = Default_Value
            Auth_Code.Enabled = True
            Auth_Code.Text = ""
            RRN.Enabled = True
            RRN.Text = ""
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = True
            NKO_Comission.Text = ""
            Payment_ID.Enabled = False
            Payment_ID.Text = Default_Value
            
        Case Is = "Платеж успешен с мобильного"
            KA_Value.Enabled = True
            KA_Value.Text = ""
            ID_Value.Enabled = True
            ID_Value.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
            Order_Number.Enabled = False
            Order_Number.Text = Default_Value
            Auth_Code.Enabled = False
            Auth_Code.Text = Default_Value
            RRN.Enabled = False
            RRN.Text = Default_Value
            Date_Value.Enabled = True
            Date_Value.Text = ""
            NKO_Comission.Enabled = True
            NKO_Comission.Text = ""
            Payment_ID.Enabled = True
            Payment_ID.Text = ""
                 
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

Private Sub CommandButton1_Click()
    If EmptyCheck() = False Then
    
            Call Add_Data(2, Ticket_Number.Text)
            Call Add_Data(3, ComboBox1.Text)
            Call Add_DV_Number(4, Ticket_Number.Text)
            Call Add_Data(5, Card_Number.Text)
            Call Add_Data(8, ID_Value.Text)
            Call Add_Data(9, KA_Value.Text)
            Call Add_Data(10, Order_Number.Text)
            Call Add_Data(11, Payment_ID.Text)
            Call Add_Data(12, Money_Value.Text)
            Call Add_Data(13, Money_Value_Kop.Text)
            Call Add_Data(14, Auth_Code.Text)
            Call Add_Data(15, RRN.Text)
            Call Add_Data(19, NKO_Comission.Text)
            
            If ComboBox1.Text = "Частичный возврат" Then
                Worksheets("Data").Cells(2, 6).Value = Replace(Trim(Date_Value.Text), " ", " в ")
                Else: Worksheets("Data").Cells(2, 6).Value = Trim(Date_Value.Text)
            End If
            If PDF_Check_Box.Value = True Then
                Worksheets("Data").Cells(2, 7).Value = "1"
                Else: Worksheets("Data").Cells(2, 7).Value = "0"
            End If
        
            MsgBox ("Успех! Далее нажми кнопку 'Сформировать подтверждение'")
            
            Unload Payment_Success
            
    Else: MsgBox ("Необходимо заполнить след. поля: Номер тикета, Вид подтверждения, Номер карты, Дата")
    End If
End Sub

