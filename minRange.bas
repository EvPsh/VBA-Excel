Function minRange(rng As String)
' Поиск минимума в произвольном диапазоне ячеек
' Пример использования min = minRange("A1:A10")
'''
Dim myRange As Range

Set myRange = Worksheets("Лист1").Range(rng)
minRange = Application.WorksheetFunction.min(myRange)

End Function