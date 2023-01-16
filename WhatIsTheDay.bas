Function WhatIsTheDay(d As Date)
'
' Ф-ция какой день недели в дате
' использование в ячейке, к примеру А1, должна быть дата 17-01-2023
' в ячейке A2 пишем = WhatIsTheDay(A1)
' результат в ячейке A2 будет написано "Вторник"
'''

Dim MyDate As Date, MyWeekDay As Integer
    
    MyWeekDay = Weekday(d)
        
    If MyWeekDay = 1 Then
        WhatIsTheDay = "Воскресенье"
    
    ElseIf MyWeekDay = 2 Then
        WhatIsTheDay = "Понедельник"
    
    ElseIf MyWeekDay = 3 Then
        WhatIsTheDay = "Вторник"
    
    ElseIf MyWeekDay = 4 Then
        WhatIsTheDay = "Среда"
        
    ElseIf MyWeekDay = 5 Then
        WhatIsTheDay = "Четверг"
        
    ElseIf MyWeekDay = 6 Then
        WhatIsTheDay = "Пятница"
        
    ElseIf MyWeekDay = 7 Then
        WhatIsTheDay = "Суббота"
        
    End If
End Function
