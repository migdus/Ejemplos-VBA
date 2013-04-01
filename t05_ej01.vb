' Autor: Miguel Dussán

' Diseñe un algoritmo que dado el número ingresado por el usuario, muestre en 
' pantalla el día de la semana al que corresponde (ejemplo: 3 es miércoles, 
' 7 es domingo). Si el usuario ingresa una opción distinta a los números de la 
' semana, deberá mostrar un mensaje de error.
Option Explicit

Sub Main()
	Dim numero_dia As Integer
	' Captura del número de día. CInt() convierte la entrada a Integer
	numero_dia = CInt(InputBox("¿Número del día de la semana?"))
	
	' Estructura de selección múltiple
	' Se examina el valor de la variable numero_dia
	
	Select Case numero_dia
		Case 1						' Si la variable tiene valor 1 imprime lunes
			MsgBox("Lunes")
		Case 2
			MsgBox("Martes")
		Case 3
			MsgBox("Miércoles")
		Case 4
			MsgBox("Jueves")
		Case 5
			MsgBox("Viernes")
		Case 6
			MsgBox("Sábado")
		Case 7
			MsgBox("Domingo")		
		Case Else ' Si no coincide ninguno de los valores con el contenido de la variable
			MsgBox("No existe un día asociado a ese número")		
	End Select
	
End Sub