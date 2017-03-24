Sub electroanalitic()
day_of_growing = Cells(18, 5).Value
day_today = ActiveSheet.Name
'*************************************************
standart = Sheets("Норматив АВІАГЕН").Select
For a = 3 To 50
If day_of_growing = Cells(a, 2) Then
myConversion = Cells(a, 9).Value
Weight = Cells(a, 7).Value
End If
Next
'*************************************************
Sheets(day_today).Select
Cells(14, 7) = Weight
Cells(16, 7) = myConversion
MsgBox ("Макрос отработал")
End Sub

