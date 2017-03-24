Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range: Set rng = [E18:F18] 'диапазон Вашей таблицы
    If Not Intersect(rng, Target) Is Nothing Then electroanalitic
End Sub
