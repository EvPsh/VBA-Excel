Sub FileInfo()
' Вывод информации о файле через FSO
' на лист Excel в ячейку A1.
'''
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set File = FSO.GetFile(Application.GetOpenFilename)
    
    Cells(1, 1).Select
    Cells(Selection.Row, 1) = File.Name
    Cells(Selection.Row, 2) = File.DateCreated
    Cells(Selection.Row, 3) = File.DateLastAccessed
    Cells(Selection.Row, 4) = File.DateLastModified
    
    Cells(Selection.Row, 5) = ParentFolder.DateCreated
    Cells(Selection.Row, 6) = ParentFolder.DateLastAccessed
    Cells(Selection.Row, 7) = ParentFolder.DateLastModified
End Sub