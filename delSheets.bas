Sub delSheets(txt as string)
' Удалить листы по маске _xxx
' пример delSheets("_txt")
' будут удалени листы с названием ZZZZZ_txt
'''

Dim txt As String
Dim j As Integer
Dim cnt As Integer

If txt = "" Then Exit Sub ' если маска не задана, выход
If Sheets.Count = 1 Then Exit Sub ' если лист в книге всего 1

cnt = 0 ' количество удалённых листов
Application.DisplayAlerts = False
	Sheets(Sheets.Count).Select
	For j = Sheets.Count To 1 Step -1
		If Right(Sheets(j).Name, Len(txt)) = txt Then Sheets(j).Delete ' если маска совпадает, удаляем лист
		cnt = cnt + 1
	Next j
Application.DisplayAlerts = True
MsgBox ("Удалено " & cnt & " листа(ов)." & chr(13) & "Всего в книге " & Sheets.Count & " листа(ов)")

End Sub