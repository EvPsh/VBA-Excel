Private Function DaysInMonth(ByVal MonthNo As Integer, Optional ByVal YearNo As Integer = 0) As Integer
' использование:
' MsgBox ("дней в месяце " & DaysInMonth(2, 2020))
' получение количества дней в месяце любого года
'''
    
    If YearNo = 0 Then YearNo = year(Date)
    DaysInMonth = Choose(MonthNo, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    
    If (MonthNo = 2) And (YearNo Mod 4 = 0) And ((YearNo Mod 100 <> 0) Or (YearNo Mod 400 = 0)) Then
        DaysInMonth = DaysInMonth + 1 ' проверка високосного года. если да, то в феврале добавляем день
    End If
    
End Function