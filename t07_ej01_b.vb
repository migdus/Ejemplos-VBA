' Autor: Miguel Dussán

' Escriba un programa que determine el promedio de calificaciones de 15 estudiantes 
' utilizando estructuras repetitivas.

Option Explicit

Sub Main()
	Dim contador, calificacion As Integer
	Dim promedio As Single
	contador = 0
	calificacion = 0
	promedio = 0
	
	While contador < 15
		calificacion = CInt(InputBox("Calificación no. "+CStr(contador+1)))
		promedio = promedio + calificacion
		contador = contador + 1
	WEnd
	
	promedio = promedio / 15
	
	MsgBox("El promedio de las calificaciones fue: "+CStr(promedio))
End Sub

