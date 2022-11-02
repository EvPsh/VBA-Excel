Private Sub Word_Out(sWord As String, cnt As Integer)
' копирование таблицы из excel в Word, 
' с последующим форматированием таблицы.
'''

Dim oWord As Object 
Dim fName As String
    
    fName = Dir("d:\temp\r_out.doc")
        If fName = "" Then
            MsgBox ("Нет файла r_out.doc")
            Exit Sub
        End If

      Set oWord = CreateObject("Word.Application") ' открывается окно word
      oWord.Visible = False ' работа без видимости окна word (быстрее и не задает лишних вопросов)
      oWord.Documents.Open ("d:\temp\r_out.doc") ' открывается файл
        Cells(cnt + 13, 6) = " Максимальное значение R2"
        Cells(cnt + 13, 11) = Cells(7, 14)
        Cells(cnt + 13, 12) = Cells(7, 15)
        Cells(cnt + 13, 13) = Cells(7, 11)
        Cells(cnt + 13, 14) = Cells(7, 12)
        Cells(cnt + 13, 15) = Cells(7, 13)
                
    Range("F" & cnt + 13 & ":J" & cnt + 13).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
                
    Range("L" & cnt + 13).Select
    Selection.Copy
    Range("M" & cnt + 13).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
                
        Range("F10:O" & cnt + 13).Select
        Selection.Copy
        
        oWord.Selection.Paste
  
    Range("F" & cnt + 13 & ":J" & cnt + 13).Select
      With Selection
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .IndentLevel = 0
          .ShrinkToFit = False
          .ReadingOrder = xlContext
          .MergeCells = False
      End With
    
      oWord.ActiveDocument.SaveAs ("d:\temp\" & sWord & ".doc")
      oWord.ActiveDocument.Close ' сохраняется предписание word
      
      oWord.Application.Quit ' закрывается word
      Set oWord = Nothing
End Sub
