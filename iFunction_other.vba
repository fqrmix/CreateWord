Option Explicit
Option Private Module

Public Function EmptyCheck() As Boolean
        If Payment_Success.Ticket_Number.Text <> "" And Payment_Success.ComboBox1.Text <> "" _
        And Payment_Success.Card_Number.Text <> "" And Payment_Success.Date_Value.Text <> "" Then
            EmptyCheck = False
            Else: EmptyCheck = True
        End If
End Function

Sub Add_Data(ByVal add_ID As Integer, ByVal add_Text As String)
    If add_Text <> "" Then
        ActiveWorkbook.Worksheets("Data").Cells(2, add_ID).Value = Trim(add_Text)
    End If
End Sub

Sub Add_DV_Number(ByVal add_ID As Integer, ByVal add_Text As String)
    If add_Text <> "" Then
        ActiveWorkbook.Worksheets("Data").Cells(2, add_ID).Value = DateValue(Now) & "/" & Trim(add_Text)
    End If
End Sub



Function ExportWord(ByVal iName As String, ByVal iVal As String) As Boolean
    With iWord.Content.Find                'With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        '.Replacement.Style = ActiveDocument.Styles("Заголовок 2 Знак")  'Возвращает или устанавливает стиль объекта.
        .Text = iName
        .Replacement.Text = iVal
        .Forward = True                     'True, если задано направление поиска "Вперёд". False - в противном случае ("Назад").
        .Wrap = wdFindContinue              'Возвращает или устанавливает константу перечисления WdFindWrap.
        .Format = False                     'True если в операцию поиска включено форматирование, False - в противном случае.
        .MatchCase = True                   'True, если в процессе поиска следует различать регистр символов.
        .MatchWholeWord = True              'True, если в процессе поиска следует искать заданный текст как отдельное слово, а не как часть другого слова.
        .MatchAllWordForms = False          'True, если требуется найти все словоформы для заданного слова.
        .MatchSoundsLike = False            'True, если требуется найти слова похожие по звучанию на заданный текст.
        .MatchWildcards = False             'True, если в процессе поиска используются регулярные выражения.

        'MatchByte                          'True, если в процессе поиска следует различать символы полной и половинной ширины.
        'ParagraphFormat                    'Возвращает или устанавливает объект ParagraphFormat.
        'Found                              'True, если в результате выполнения поиска было найдено соответствие.
        'Font                               'Возвращает или устанавливает объект Font задающий форматирование шрифта.
        
        '.Execute Replace:=wdReplaceAll
        .Execute
        
        If .Found Then  'проверяем, найдена ли Закладка в документе Word
            ExportWord = True           'закладка найдена
            .Execute Replace:=wdReplaceAll
        Else
            ExportWord = False          'закладка НЕ найдена
        End If
    End With
End Function


Sub FolderCreateDel(ByVal iPath As String)
    Dim BasePath As String
    On Error Resume Next
    Kill iPath & "*.docx"
    MkDir (iPath)
End Sub


