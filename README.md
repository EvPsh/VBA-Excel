# VBA-Excel

delSheets.bas - Sub delSheets(txt as string). Удалить листы по маске. В названии файла должна быть фраза txt.

getDirs.bas - Function Get_DirS(path As String, Mask As String). Функция выбора в массив файлов по маске, с примером использования.

iColor.bas - Sub intcolor(). Подсчёт значений в 1ом столбце, выделенном жёлтым цветом

numTextBox.bas - Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger). Ограничение ввода символов в поле TextBox1 (только цифры и '.')

oExcel.bas - sub oExcel(). Пример, показывающий как отрыть файл и передать его в объект oExcel.

oWord.bas - Private Sub Word_Out(sWord As String, cnt As Integer). Копирование таблицы из excel в Word, с последующим форматированием таблицы.

shNamesOut.bas - Private Sub SheetsNameOut(). Вывод имён листов в новую книгу.

xlsxToxls.bas - Private Sub XlsxToXls(FullName As String). Сохранение файлов из XLSX в формат XLS 97-2003.