Attribute VB_Name = "M�dulo1"
Sub teste()

Dim i As Integer

Range("D7").Value = 1

For i = 1 To 4
Cells(i, 1).Value = "0"
Next i

For i = 5 To 8
Cells(i, 1).Value = "1"
Next i


End Sub
