Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
' Ограничение ввода символов в поле TextBox1
' вводятся только числа и '.'
'''

 If KeyAscii < 46 Or KeyAscii = 47 Or KeyAscii > 57 Then KeyAscii = 0
End Sub