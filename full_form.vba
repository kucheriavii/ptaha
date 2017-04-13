Function Sh_Exist(wb As Workbook, sName) As Boolean
    Dim wsSh As Worksheet
    On Error Resume Next
    Set wsSh = wb.Sheets(sName)
    Sh_Exist = Not wsSh Is Nothing
End Function
Sub copier() 'копіює дані стягнуті з 1С. ЧОМУ ВІДРАЗУ НЕ ПРАЦЮЮ В ФАЙЛІ ЗБЕРЕЖЕНОМУ З 1С? - ТА ТОМУ ЩО ТАК Я ЗАХОТІВ, І КОЛИ З 1С КОПІЮЄШ ФОРМАТ КНИГИ ХЄРИТЬСЯ, І КОПІЮВАТИ НАДІЙНІШЕ - НІХТО НЕ ВЛІЗЕ В ФАЙЛ ОРИГІНАЛУ
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    copybookname = "БАЗА_ДАНИХ" & ActiveWorkbook.Name
    copybookname = Left(copybookname, Len(copybookname) - 1)
    bookname = ActiveWorkbook.Name
    'MsgBox (copybookname) 'for debuging
    adr = "D:\Чебатурочка\Виробнича\Аналітичний відділ\ПЕРЕДАНІ ФОРМИ ОБЛІКУ\БАЗА ДАНИХ\" & copybookname
    Workbooks.Open Filename:=adr
    Windows(copybookname).Activate
    Cells.Select
    Range("E1000").Activate
    Selection.Copy
    Windows(bookname).Activate
    Sheets(1).Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    Windows(copybookname).Activate
    Workbooks(copybookname).Close SaveChanges:=False
    Windows(bookname).Activate
End Sub
Sub build_summary_file()
'ЯКЩО В ТЕБЕ ПОХЄРИВСЯ МАКРОС - НЄХРЄН ЧЕКАТИ ПОКИ Я ПРИЙДУ І ПОФІКШУ! ПЕРША ПРИЧИНА ПОЛОМКИ - СТОРІНКА, НА ЯКІЙ ПОВИНЕН ЗАПУСКАТИСЯ МАКРОС
'ДА...ДА...ДА ТА САМА В ЯКІЙ В ТЕБЕ ЗБЕРІГАЮТЬСЯ ВСІ ДАНІ В КУЧІ, ПОВИННА БУТИ ПЕРШОЮ В СПИСКУ ВСІХ СТОРІНОК, НЕ ДРУГОЮ НЕ ТРЕТЬОЮ, А ПЕРШОЮ
'ЧОГО? -БО, Я ЗВЕРТАЮСЯ ДО НЕЇ ЯК shits(1) - А ЦЕ ЗНАЧИТЬ ПЕРША СТОРІНКА!!!
'Друга причина поломки в рядку "If Left(houseEnd, 7) = "Пташник" Or IsEmpty(Cells(j, 1)) Then" іноді, коли в 1с є порожні значення ячейки і значення слыд замынити на "If Left(houseEnd, 7) = "Пташник" Or IsEmpty(Cells(j, 1)) Then"
Sheets(1).Activate
For i = 1 To 1000 'перебираємо в циклі перший рядок з метою знайти ячейку де починається кожен новий пташник
    building = Cells(i, 1).Value
    If Left(building, 7) = "Пташник" Then 'Якщо значення в ячейці починається на "Пташник" - значить це початок нового пташника
    
    If Not Sh_Exist(ActiveWorkbook, CStr(building)) Then 'Цю функцію я вкрав на stackoverflow, вона приймає значення книги і листа і якщо листа з потрібним іменем немає
        ActiveWorkbook.Sheets.Add(, ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)).Name = CStr(building) '- то створює його і запихає в кінець
        Sheets(1).Activate 'тут починаємо процедуру створення шапки для кожного листа
        Range("A7:CA10").Copy 'копіюємо як лохи шапку з великим запасом, моло чого ще придумають в 1с
        Sheets(CStr(building)).Activate 'активуємо щойностворену сторінку
        Cells(1, 1).Activate 'шукаємо початок на сторінці - ячейку А1 і виділяємо її (аналог кліка мишки)
        ActiveSheet.Paste 'копіюємо виділену раніше шапку (див. копіюємо як лохи)
        Sheets(1).Activate 'повертаємось на головну сторінку з якої виконується макрос вона завжди повинна бути першою
    End If
        
        
        houseStart = i 'початок пташника
       ' MsgBox (building & " починається на " & i + 1 & " ,1")
        
        For j = i + 1 To i + 100 'шукаємо закінчення пташника
                houseEnd = Cells(j, 1) 'від того значення де було знайдено початок перебираємо вниз 100 значень (42 дня вирощування + запас 48 днів)
                If Left(houseEnd, 7) = "Пташник" Or IsEmpty(Cells(j, 4)) Then 'якщо знайшли значення пташник то записуємо його
                houseEnd = j 'записали
               ' MsgBox (building & " закінчується на " & j - 1 & " ,1") 'ДЛЯ ДЕБАЖІННЯ (ПРОВІРКИ РОБОТИ МАКРОСА)
                Range(Cells(i, 1), Cells(houseEnd - 1, 100)).Copy 'Виділяємо ту частину яка нас цікавить: Cells(i, 1) - перша ячейка, Cells(houseEnd - 1, 100) - остання ячейка, houseEnd - 1 - для того що houseEnd - це ячейка в якій вже закінчився пташник (границя), а для виділення нам потрібно лише останнє значення. Границя (для допитливих) - це імя наступного пташника
                Sheets(CStr(building)).Activate 'Активуємо потрібний лист
                Cells(5, 1).Activate 'вибираємо 5 строчку (рядок), 4 попередні виділені під шапку
                ActiveSheet.Paste 'запихаємо дані в лист
                'Cells(1, 1) = building
                Sheets(1).Activate 'повертаємось на лист з макросом
                GoTo nxt
            End If
        Next
    End If
nxt:
Next
End Sub
Sub clearNegativeDays()
    cPages = ThisWorkbook.Worksheets.Count
        Sheets(1).Activate
        For xzz = 1000 To 1 Step -1
        mstr = Cells(xzz, 4).Value
        c = InStr(1, mstr, "-")
        If c > 0 Then
            Rows(xzz).Delete
        End If
        Next
End Sub
Sub DoAll()
copier
clearNegativeDays
build_summary_file
Sheets(1).Activate
End Sub





