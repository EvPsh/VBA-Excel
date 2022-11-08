Sub RenameFiles()
' Переименовывание файлов *.doc
' в первом столбце Excel должны быть необходимые имена
' в выбираемой папке - такое же количество файлов, сколько имён в столбце 1 excel
'''
    Dim OldName As String, NewName As String, MyPath As String
    Dim fName As String, i as integer
    
    MyPath = InputBox("путь к папке") & "\"
    fName = Dir(MyPath & "*.doc")
    
    If fName = "" Then
        MsgBox ("Нет нужных файлов")
        Exit Sub
    End If
        
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, 1) <> "" Then
            OldName = MyPath & fName
            NewName = MyPath & Cells(i, 1) & ".doc" 
            Name OldName As NewName
            fName = Dir
        End If
    Next

MsgBox "Rename Complete", vbOKOnly

End Sub
