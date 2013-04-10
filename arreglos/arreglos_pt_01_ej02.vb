' Autor: Miguel Dussán

Option Explicit

Sub main()

Dim i As Integer ' Variable contadora de ciclo

' Declaración del arreglo de 10 elementos de tipo Integer
Dim a(10) As Integer

'Inicialización del arreglo
'El valor entre paréntesis es el índice en el arreglo
a(0) = 3
a(1) = 56
a(2) = 4
a(3) = 2
a(4) = 30
a(5) = 12
a(6) = 6
a(7) = 4
a(8) = 73
a(9) = 49

Cells(1, 1) = "Posición"
Cells(1, 2) = "Valor"

For i = 0 To 10
    Cells(i + 2, 1) = i
    Cells(i + 2, 2) = a(i)
Next

End Sub