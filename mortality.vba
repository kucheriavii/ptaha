Sub Рознести()

n = 0
k = 2
p = 3
r = 2
For x = 3 To 15
For i = 10 To 59
k = k + 1
If (Sheets("Лист1").Cells(k, r).Value > 0) Then
Cells(i, p) = Sheets("Лист1").Cells(k, r)
End If
If (Sheets("Лист1").Cells(k, r + 1).Value > 0) Then
Cells(i, p + 1) = Sheets("Лист1").Cells(k, r + 1)
End If
Next
p = p + 7
r = r + 2
k = 2
Next
End Sub
Sub mortality()
Application.DisplayAlerts = False
user = Environ("USERNAME") 'Команда проверяет имя пользователя
Dim users_with_access(1 To 10) As String 'объявление массива с пользователями у которых будет доступ к работе макроса
'Блок пользователей с доступом к макросу
'****************************************************
'откроем файл assets.xlsx (в этом файле сохранены все юзеры с доступом к макросу).
Application.Workbooks.Open ("D:\Чебатурочка\Виробнича\Аналітичний відділ\ПЕРЕДАНІ ФОРМИ ОБЛІКУ\БАЗА ДАНИХ\assets.xlsx")
'переберем записи чтобы выбрать имена привелигиованных пользователей
For j = 1 To 10
'присвоем именна пользователей в массив
users_with_access(j) = Cells(j, 1).Value
Next j
'закроем файл с именами
ActiveWindow.Close
'****************************************************
'доступ рассщитан на 10 пользователей.
For i = 1 To 10
    'если имя пользователя в списке - макрос запустится
    If users_with_access(i) = user Then
        'MsgBox (users_with_access(i) & " is a correct user. Your are welcome!")
        Рознести
        GoTo mac
    End If
    'если счетчик дошел до 10, а имя пользователя не было знайдено - макрос не запустится
    If i = 10 And user_with_access <> user Then
        MsgBox ("Дорогой(-я) " & user & " Доступ к работе макроса вам закрыт. Если желаете облегчить себе жизнь обращайтесь в диспетчерскую службу. Удачи вам!")
        GoTo finish
    End If
    Next i
mac:
    Рознести
    
finish:
    MsgBox ("Спасибо чо пользуетесь нашим продуктом.")
End Sub



