# VBA-Excel

RenameFiles.bas - Sub RenameFiles(). Переименовывание файлов *.doc именами из первого столбца в Excel.

delSheets.bas - Sub delSheets(txt as string). Удалить листы по маске. В названии листа должна быть фраза txt.

findInColumn.bas - Sub poisk(). Поиск данных в диапазоне ячеек листа Excel.

getDirs.bas - Function Get_DirS(path As String, Mask As String). Функция выбора в массив файлов по маске, с примером использования.

lName.bas - Private Function LName(). Функция переформатирования даты из xx.yy.zzzz в zzzz-yy-xx.

iColor.bas - Sub intcolor(). Подсчёт значений в 1ом столбце, выделенном жёлтым цветом.

minRange.bas - Function minRange(rng As String). Поиск минимума в произвольном диапазоне ячеек.

numTextBox.bas - Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger). Ограничение ввода символов в поле TextBox1 (только цифры и '.')

oExcel.bas - sub oExcel(). Пример, показывающий как отрыть файл и передать его в объект oExcel.

oWord.bas - Private Sub Word_Out(sWord As String, cnt As Integer). Копирование таблицы из excel в Word, с последующим форматированием таблицы.

shNamesOut.bas - Private Sub SheetsNameOut(). Вывод имён листов в новую книгу.

xlsxToxls.bas - Private Sub XlsxToXls(FullName As String). Сохранение файлов из XLSX в формат XLS 97-2003.
