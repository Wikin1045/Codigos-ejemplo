Attribute VB_Name = "formato_fecha"

Public rango As String
Private Sub Worksheet_Change(ByVal Target As Range)
Dim x As String
Dim m As String
rango = ActiveCell.Address
m = 1
For i = 1 To 4
Select Case i
Case Is = 1: x = "o"
Case Is = 2: x = "p"
Case Is = 3: x = "r"
Case Is = 4: x = "s"
End Select
x = x & m
Range(x).Select
Do While ActiveCell <> Empty
ActiveCell.NumberFormat = "[$-es-ES]dd-mmm-yy          h:mm "" h "" ;@"
ActiveCell.Offset(1, 0).Select
Loop
Next
Range(rango).Select
End Sub


