Function Get_DirS(path As String, Mask As String) 
' Функция выбора в массив файлов по маске
' path - путь к файлу, пример 'd:\tmp\'
' Mask - маска для выбора файлов, пример '*.*'
'''
Dim a() As String, D As String, U As Long

D = Dir(path & Mask) ' vbDirectory)
While D <> ""
    ReDim Preserve a(U)
    'a(U) = path & D
    a(U) = D
    U = U + 1
    D = Dir
Wend
Get_DirS = a
End Function

Sub test()
' Ф-ция для показа работы Get_DirS
' вывод файлов *.doc из папки d:\tmp
'''
Dim a() As String, i As Integer

a() = Get_DirS("d:\Tmp\", "*.doc") ' путь, маска
For i = 0 To UBound(a())
    'Debug.Print a(i) ' возврат файлов *.doc 
    Cells(i + 1, 1) = a(i) ' вывод названий файлов на лист Excel
Next i
End Sub