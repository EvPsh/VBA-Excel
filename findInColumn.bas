Sub poisk()
' Поиск в диапазоне заданных значений
'''

Dim с As String
Dim m As String

m = InputBox("Что ищем", "Поиск")
With Worksheets("00").Range("a1:a100") ' выбор диапазона поиска
    Set c = .Find(m) ' что искать
    Cells(1, 1) = c.Row
End With

End Sub