Private Function SbVs(d As Integer, ByVal m As Integer, y As Integer) As Boolean
' 
' Ф-ция получения сб и вс. Если дата = сб или = вс, то ф-ция = true
' на вход:
' d - день
' m - месяц
' y - год
' на выходе true/false
'''

Dim MyDate As Date, MyWeekDay As Integer
    
    If m > 12 Then
        MsgBox ("Месяц > 12, выход")
        SbVs = False
        Exit Function
    End If
    
    MyDate = CDate(d & "/" & m & "/" & y) ' день/месяц/год
    MyWeekDay = Weekday(MyDate)
    
    If MyWeekDay = 7 Or MyWeekDay = 1 Then 
    ' день недели начинается с вс ( = 1), заканчивается сб( = 7)
        SbVs = True
    Else
        SbVs = False
    End If
    
End Function
