sub oExcel()
' пример, показывающий как отрыть файл и передать его в объект oExcel
'''

Dim oExcel As Object ' объект книги Excel
Dim MyDial As FileDialog, xFileDial As String ' для диалогового окна

	'Application.ScreenUpdating = False ' 
	    
	'-------------------------------------- выбор документа -------------------------------------------
	Set MyDial = Application.FileDialog(msoFileDialogOpen)

	MyDial.AllowMultiSelect = False
	MyDial.Filters.Clear
	MyDial.Filters.Add "Excel", "*.xls*"
	MyDial.Title = "ВЫБОР ФАЙЛА EXCEL"
	MyDial.Show
	 
	If MyDial.SelectedItems.Count > 0 Then
	    xFileDial = MyDial.SelectedItems(1)
	Else
	    Exit Sub
	End If
	' -------------------------------------------------------------------------------------------------
	Set oExcel = CreateObject("Excel.Application") ' открывается окно Excel

	With oExcel
		.Visible = True ' видимость окна
		.Workbooks.Open fileName:=xFileDial ' открываем файл с данными в объект oExcel

		    MsgBox ("Кол-во листов в книге" & .ActiveWorkbook.Sheets.Count), , "" ' для примера выводим количество листов книги

		'.ActiveWindow.Close ' закрываем книгу Excel
		'.Application.Quit	' выходим из Excel
	End With

	Set oExcel = Nothing 	' закрываем объект

	Application.ScreenUpdating = True
end sub