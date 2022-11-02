Private Function LName()
' 
' функция переформатирования даты из xx.yy.zzzz в zzzz-yy-xx
' #форматирование дат #обработка дат #дата в название листа
' использование Sheets("Лист1").Name = LName
'''

Dim a As String
Dim a1, a2, a3 As String

a = Date
a1 = Mid(a, 1, 2)
a2 = Mid(a, 4, 2)
a3 = Mid(a, 7, 4)
a = a3 & "-" & a2 & "-" & a1

LName = a
End Function