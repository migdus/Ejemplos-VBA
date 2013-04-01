' Autor: Miguel Duss�n

' Dise�e un algoritmo que dado el n�mero ingresado por el usuario, muestre en 
' pantalla el d�a de la semana al que corresponde (ejemplo: 3 es mi�rcoles, 
' 7 es domingo). Si el usuario ingresa una opci�n distinta a los n�meros de la 
' semana, deber� mostrar un mensaje de error.
Option Explicit

Sub Main()
	Dim numero_dia As Integer
	' Captura del n�mero de d�a. CInt() convierte la entrada a Integer
	numero_dia = CInt(InputBox("�N�mero del d�a de la semana?"))
	
	' Estructura de selecci�n m�ltiple
	' Se examina el valor de la variable numero_dia
	
	Select Case numero_dia
		Case 1						' Si la variable tiene valor 1 imprime lunes
			MsgBox("Lunes")
		Case 2
			MsgBox("Martes")
		Case 3
			MsgBox("Mi�rcoles")
		Case 4
			MsgBox("Jueves")
		Case 5
			MsgBox("Viernes")
		Case 6
			MsgBox("S�bado")
		Case 7
			MsgBox("Domingo")		
		Case Else ' Si no coincide ninguno de los valores con el contenido de la variable
			MsgBox("No existe un d�a asociado a ese n�mero")		
	End Select
	
End Sub