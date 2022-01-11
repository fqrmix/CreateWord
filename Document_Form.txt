Private Sub UserForm_Initialize()

    With Me.ComboBox1
        .ControlTipText = "Выберите значение из списка"
        .List = Array("Подтверждение остатка - ИД", "Подтверждение остатка - УИ", "Подтверждение остатка + карта - ИД" _
        , "Подтверждение остатка + карта - УИ", "Подтверждение остатка на старую дату - УИ", "Подтверждение остатка на старую дату - ИД" _
        , "Подтверждение остатка с датой открытия кошелька - УИ", "Подтверждение остатка с датой открытия кошелька - ИД", "Подтверждение остатка с ПД - ИД" _
        , "Подтверждение остатка с ПД - УИ", "Подтверждение о закрытии кошелька - ИД", "Подтверждение о закрытии кошелька - УИ" _
        , "Шаблон подтверждение закрытой карты - ИД", "Шаблон подтверждение закрытой карты - УИ")
    End With

End Sub

Function Current_Date()

    Dim m$
    
    m = LCase$(MonthName$(Month(Now)))
    
    If Right$(m, 1) = "т" Then
        m = m & "а"
        Else: m = Left$(m, Len(m) - 1) & "я"
        
    End If
        
    Current_Date = Format(Now, "DD """ & m & " ""YYYY")
End Function

Function Current_Time()

    Current_Time = Format(Now, "hh:mm")
    
End Function

Private Sub ComboBox1_Change()

Dim Default_Value As String
Default_Value = "Заполнение не требуется"
        
    Select Case ComboBox1.Text
    
        Case Is = "Подтверждение остатка - ИД"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            PassportID.Text = ""
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True
            
        Case Is = "Подтверждение остатка - УИ"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True
            
        Case Is = "Подтверждение остатка + карта - ИД"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = ""
            YooCard.Enabled = True
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True
            
        Case Is = "Подтверждение остатка + карта - УИ"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = ""
            YooCard.Enabled = True
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True
            
        Case Is = "Подтверждение остатка на старую дату - УИ"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = ""
            Past_Money_Value.Enabled = True
            Past_Date.Text = ""
            Past_Date.Enabled = True
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True
            
        Case Is = "Подтверждение остатка на старую дату - ИД"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = ""
            Past_Money_Value.Enabled = True
            Past_Date.Text = ""
            Past_Date.Enabled = True
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True
            
        Case Is = "Подтверждение остатка с датой открытия кошелька - УИ"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = ""
            Wallet_Date.Enabled = True
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True

        Case Is = "Подтверждение остатка с датой открытия кошелька - УИ"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = ""
            Wallet_Date.Enabled = True
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True
            
        Case Is = "Подтверждение остатка с ПД - УИ"
            Birth_Date.Text = ""
            Birth_Date.Enabled = True
            Given_Date.Text = ""
            WhoIs.Enabled = True
            Document_Type.Text = ""
            Document_Type.Enabled = True
            Given_Date.Text = ""
            Given_Date.Enabled = True
            WhoIs.Text = ""
            WhoIs.Enabled = True
            Document_Code.Text = ""
            Document_Code.Enabled = True
            Birth_Place.Text = ""
            Birth_Place.Enabled = True
            Reg_Place.Text = ""
            Reg_Place.Enabled = True
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True
            
        Case Is = "Подтверждение остатка с ПД - ИД"
            Birth_Date.Text = ""
            Birth_Date.Enabled = True
            Given_Date.Text = ""
            WhoIs.Enabled = True
            Document_Type.Text = ""
            Document_Type.Enabled = True
            Given_Date.Text = ""
            Given_Date.Enabled = True
            WhoIs.Text = ""
            WhoIs.Enabled = True
            Document_Code.Text = ""
            Document_Code.Enabled = True
            Birth_Place.Text = ""
            Birth_Place.Enabled = True
            Reg_Place.Text = ""
            Reg_Place.Enabled = True
            Closing_Date.Text = Default_Value
            Closing_Date.Enabled = False
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = ""
            Current_Money_Value.Enabled = True
            
        Case Is = "Подтверждение о закрытии кошелька - ИД"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = ""
            Closing_Date.Enabled = True
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = Default_Value
            Current_Money_Value.Enabled = False
            
        Case Is = "Подтверждение о закрытии кошелька - УИ"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = ""
            Closing_Date.Enabled = True
            Closed_YooCard.Text = Default_Value
            Closed_YooCard.Enabled = False
            YooCard.Text = Default_Value
            YooCard.Enabled = False
            Wallet_Date.Text = Default_Value
            Wallet_Date.Enabled = False
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = ""
            PassportID.Enabled = True
            Country.Text = ""
            Country.Enabled = True
            Current_Money_Value.Text = Default_Value
            Current_Money_Value.Enabled = False
            
        Case Is = "Шаблон подтверждение закрытой карты - ИД"
            Birth_Date.Text = Default_Value
            Birth_Date.Enabled = False
            Given_Date.Text = Default_Value
            WhoIs.Enabled = False
            Document_Type.Text = Default_Value
            Document_Type.Enabled = False
            Given_Date.Text = Default_Value
            Given_Date.Enabled = False
            WhoIs.Text = Default_Value
            WhoIs.Enabled = False
            Document_Code.Text = Default_Value
            Document_Code.Enabled = False
            Birth_Place.Text = Default_Value
            Birth_Place.Enabled = False
            Reg_Place.Text = Default_Value
            Reg_Place.Enabled = False
            Closing_Date.Text = ""
            Closing_Date.Enabled = True
            Closed_YooCard.Text = ""
            Closed_YooCard.Enabled = True
            YooCard.Text = ""
            YooCard.Enabled = True
            Wallet_Date.Text = ""
            Wallet_Date.Enabled = True
            Past_Money_Value.Text = Default_Value
            Past_Money_Value.Enabled = False
            Past_Date.Text = Default_Value
            Past_Date.Enabled = False
            PassportID.Text = Default_Value
            PassportID.Enabled = False
            Country.Text = Default_Value
            Country.Enabled = False
            Current_Money_Value.Text = Default_Value
            Current_Money_Value.Enabled = False
            
            
            
    End Select
End Sub

Private Sub CommandButton1_Click()
    
            Call Add_Data(2, Ticket_Number.Text)
            Call Add_Data(3, ComboBox1.Text)
            Call Add_DV_Number(4, Ticket_Number.Text)
            Call Add_Data(5, Card_Number.Text)
            Call Add_Data(20, Country.Text)
            Call Add_Data(21, PassportID.Text)
            Call Add_Data(22, FullName.Text)
            Call Add_Data(23, Current_Money_Value.Text)
            Call Add_Data(24, Past_Money_Value.Text)
            Call Add_Data(25, Past_Date.Text)
            Call Add_Data(26, YooCard.Text)
            Call Add_Data(27, Current_Date())
            Call Add_Data(28, Current_Time())
            If Current_Money_Value.Enabled = True Then
                If Current_Money_Value.Text = "0,00" Then
                        Worksheets("Data").Cells(2, 29).Value = "Ноль рублей ноль копеек"
                    Else:
                    Call Add_Data(29, SumToWord(CDbl(Current_Money_Value.Text)))
                End If
            End If
            
            If Past_Money_Value.Enabled = True Then
                Call Add_Data(30, SumToWord(CDbl(Past_Money_Value.Text)))
            End If
            Call Add_Data(31, Wallet_Date.Text)
            If ComboBox1.Text = "Подтверждение остатка с ПД - ИД" Or _
                ComboBox1.Text = "Подтверждение остатка с ПД - УИ" Then
                    Call Add_Data(32, Birth_Date.Text)
                    Call Add_Data(33, Document_Type.Text)
                    Call Add_Data(34, Given_Date.Text)
                    Call Add_Data(35, WhoIs.Text)
                    Call Add_Data(36, Document_Code.Text)
                    Call Add_Data(37, Birth_Place.Text)
                    Call Add_Data(38, Reg_Place.Text)
            End If
            Call Add_Data(39, Closing_Date.Text)
            Call Add_Data(40, Closed_YooCard.Text)
            
            If PDF_Check_Box.Value = True Then
                Worksheets("Data").Cells(2, 7).Value = "1"
                Else: Worksheets("Data").Cells(2, 7).Value = "0"
            End If
            
            MsgBox ("Успех! Далее нажми кнопку 'Сформировать подтверждение'")
            
            Unload Official_DocForm
End Sub
