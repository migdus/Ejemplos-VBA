' Autor: Miguel Dussán

Option Explicit

Sub main()
' Declaración del arreglo de 10 elementos de tipo Integer
Dim a(10) As Integer

Dim i As Integer ' Variable contadora de ciclo

Dim res As String ' Variable acumuladora para salida
res = ""

' Inicialización del arreglo
' Se debe tomar cada una una de las posiciones y asignarles un valor

For i = 0 To 10
	a(i) = 0
Next

' Mostrar el valor de la información almacenada en el arreglo
res = "Posición" + vbTab + "Valor" + vbCrLf

For i = 0 To 10
	res = res + CStr(i) + vbTab + vbTab + vbTab + vbTab + CStr(a(i)) + vbCrLf
Next

MsgBox (res)

' Esta información también se puede escribir en las celdas de la hoja de cálculo
' Cells(x,y) = "cadena" escribe en las celdas x,y (tomando la celda A1 como 1,1)
' la cadena
Cells(1, 1) = "Posición"
Cells(1, 2) = "Valor"

For i = 0 To 10
'A las coordenadas en x se le suman 2 porque i arranca en 0, y Cells(x,y)
'Solo admite valores desde 1. Además así no se sobreescriben las cadenas 
'Posición y Valor, que van como encabezados de las columnas'
    Cells(i + 2, 1) = i
    Cells(i + 2, 2) = a(i)
Next

End Sub
