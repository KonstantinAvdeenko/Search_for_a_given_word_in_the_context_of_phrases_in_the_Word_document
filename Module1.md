Private Sub CommandButton1_Click()
'находит и показывает фразы внутри документа word, где находится нужное пользователю слово или словосочетание'
найдено = 0
'проверяет задал ли пользователь слово для поиска во фразах'
If (TextBox1.Text <> "") And (TextBox2.Text = "") And (TextBox3.Text = "") Then
    For i = 1 To ActiveDocument.Sentences.Count 'Для поиска берется полностью весь текст'
    'Выполнение чувствительной к регистру операции поиска слов и возврат результата логического да если поиск был удачен'
       If ActiveDocument.Sentences(i).Find.Execute(FindText:=TextBox1.Text, MatchCase:=True, MatchWholeWord:=True) Then
       'Добавление нужной фразы'
найдено = найдено + 1
ActiveDocument.Sentences(i).Select 'Выборка нужного контекста фразы'
MsgBox ("Контекст найден")
End If
    Next i
    Else: MsgBox ("Вы ничего не ввели для поиска")
End If
If (TextBox1.Text <> "") And (TextBox2.Text <> "") And (TextBox3.Text = "") Then
    For i = 1 To ActiveDocument.Sentences.Count
    'Выполнение операции поиска и сравнения слов, фраз, отсеивание лишнего контекста, 
     'возврат результата логического да если поиск был удачен'
       If ActiveDocument.Sentences(i).Find.Execute(FindText:=TextBox1.Text, MatchCase:=True, MatchWholeWord:=True) And ActiveDocument.Sentences(i).Find.Execute(FindText:=TextBox2.Text, MatchCase:=True, MatchWholeWord:=True) Then
найдено = найдено + 1
ActiveDocument.Sentences(i).Select
MsgBox ("Более точный контекст найден")
End If
    Next i
End If
If (TextBox1.Text <> "") And (TextBox2.Text <> "") And (TextBox3.Text <> "") Then
    For i = 1 To ActiveDocument.Sentences.Count
    'Выполнение операции поиска и сравнения слов, фраз, отсеивание лишнего контекста, 
      'возврат результата логического да если поиск был удачен'
       If ActiveDocument.Sentences(i).Find.Execute(FindText:=TextBox1.Text, MatchCase:=True, MatchWholeWord:=True) And ActiveDocument.Sentences(i).Find.Execute(FindText:=TextBox2.Text, MatchCase:=True, MatchWholeWord:=True) And ActiveDocument.Sentences(i).Find.Execute(FindText:=TextBox3.Text, MatchCase:=True, MatchWholeWord:=True) Then
найдено = найдено + 1
ActiveDocument.Sentences(i).Select
MsgBox ("Еще более точный контекст найден")
End If
    Next i
End If
If найдено > 0 Then MsgBox ("найдено" & найдено & "вхождений текста")
Selection.Collapse 'Сброс выборки'
End Sub

Private Sub CommandButton2_Click()
'Сброс выборки и закрытие формы'
Selection.Collapse
UserForm1.Hide
Set UserForm1 = Nothing
End Sub