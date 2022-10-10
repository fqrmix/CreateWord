Option Explicit

'Без копеек (1), с копейками (0)
'Копейки прописью(1), числом(0)
'Начинать прописью(0), заглавной(1)

Function SumToWord(Sum As Double, _
    Optional Without_Kop As Boolean = False, _
    Optional Write_Kop As Boolean = True, _
    Optional Up_Start As Boolean = True) As String
    
'Функция для написания суммы прописью
 Dim ed, des, sot, ten, razr, dec
 Dim i As Integer, str As String, s As String
 Dim intPart As String, frPart As String
 Dim mlnEnd, tscEnd, razrEnd, rub, cop
        dec = Array(" ", "одна ", "две ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
        ed = Array(" ", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
        ten = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")
        des = Array("", "", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
        sot = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
        razr = Array("", "тысяч", "миллион", "миллиард")
        mlnEnd = Array("ов ", " ", "а ", "а ", "а ", "ов ", "ов ", "ов ", "ов ", "ов ")
        tscEnd = Array(" ", "а ", "и ", "и ", "и ", " ", " ", " ", " ", " ")
        razrEnd = Array(mlnEnd, mlnEnd, tscEnd, "")
        rub = Array("рублей", "рубль", "рубля", "рубля", "рубля", "рублей", "рублей", "рублей", "рублей", "рублей")
        cop = Array("копеек", "копейка", "копейки", "копейки", "копейки", "копеек", "копеек", "копеек", "копеек", "копеек")
        
        
 If Sum >= 1000000000000# Or Sum < 0 Then SumToWord = CVErr(xlErrValue): Exit Function

 If Round(Sum, 2) >= 1 Then
 
    intPart = Left$(Format(Sum, "000000000000.00"), 12)
    
    For i = 0 To 3
        s = Mid$(intPart, i * 3 + 1, 3)
        If s <> "000" Then
            str = str & sot(CInt(Left$(s, 1)))
                If Mid$(s, 2, 1) = "1" Then
                    str = str & ten(CInt(Right$(s, 1)))
                Else
                    str = str & des(CInt(Mid$(s, 2, 1))) & IIf(i = 2, dec(CInt(Right$(s, 1))), ed(CInt(Right$(s, 1))))
                End If
                

                On Error Resume Next
                
                str = str & IIf(Mid$(s, 2, 1) = "1", razr(3 - i) & razrEnd(i)(0), _
                razr(3 - i) & razrEnd(i)(CInt(Right$(s, 1))))
                
                On Error GoTo 0
            End If
            Next i
            str = str & IIf(Mid$(s, 2, 1) = "1", rub(0), rub(CInt(Right$(s, 1))))
        End If
        
 SumToWord = str
 ''''''''''''''''''
 If Without_Kop = False Then
    frPart = Right$(Format(Sum, "0.00"), 2)
            If Write_Kop Then
                frPart = IIf(Left$(frPart, 1) = "1", ten(CInt(Right$(frPart, 1))) & cop(0), _
                des(CInt(Left$(frPart, 1))) & dec(CInt(Right$(frPart, 1))) & cop(CInt(Right$(frPart, 1))))
            Else
                frPart = IIf(Left$(frPart, 1) = "1", frPart & " " & cop(0), frPart & " " & cop(CInt(Right$(frPart, 1))))
            End If
        SumToWord = str & " " & frPart
 End If
 
 ''''''''''''''''''
 
If Up_Start Then Mid$(SumToWord, 1, 1) = UCase(Mid$(SumToWord, 1, 1))
 
End Function
