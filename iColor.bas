Sub intcolor()
' Подсчёт значений в 1ом столбце, выделенном жёлтым цветом
'''
Dim i As Integer
Dim b As Integer ' номер колонки
Dim res As Integer ' результат
Dim endColumn As integer ' последняя заполненная ячейка в 1ом столбце

b = 1 ' номер колонки
endColumn = Cells(Rows.Count, b).End(xlUp).Row ' последнее значение в колонке b

res = 0 ' сумма всех выделенных жёлтым ячеек

For i = 1 To b
    If Cells(i, b).Interior.Color = RGB(255, 255, 0) Then
        res = res + Cells(i, b).Value
        
    End If
Next i
MsgBox ("Сумма в выделенных ячейках = " & res),,""
End Sub