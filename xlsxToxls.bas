Private Sub XlsxToXls(FullName As String)
' Сохранение файла из XLSX в формат XLS 97-2003
'''
	Application.Workbooks.Open (FullName)
	Application.ScreenUpdating = False
    
    If ActiveWorkbook.RemovePersonalInformation Then
        ActiveWorkbook.RemovePersonalInformation = False
    End If
    
    Application.DisplayAlerts = False
    
    FullName = Left(FullName, Len(FullName) - 5) ' обрезка xlsx
    ActiveWorkbook.SaveAs Filename:=FullName & ".xls", FileFormat:=xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWorkbook.Close

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub