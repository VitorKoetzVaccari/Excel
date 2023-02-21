Attribute VB_Name = "Módulo1"
Sub teste()

Dim i As Integer

entradas = Application.InputBox("Quantas entradas?")

For i = 1 To 8
    If i Mod 2 = 1 Then
        Cells(i, 3).Value = 0
        Cells(i, 2).Value = 0
        Cells(i + 1, 2).Value = 0
    Else
        Cells(i, 3).Value = 1
    End If
    
Next i
    
End Sub
