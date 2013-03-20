'Autor: Miguel Dussán

' Diseñe un programa que emule un control de acceso. El programa debe preguntar al
' usuario por una contraseña alfanumérica (es decir, que puede contener números y 
' letras) y verificar si la contraseña es correcta. El usuario dispondrá de tres 
' intentos para acertar la contraseña. Si el usuario escribe la contraseña de 
' forma correcta, el programa debe mostrar un mensaje que diga: Acceso concedido; 
' de lo contrario mostrará el mensaje: Acceso denegado. La contraseña que estará 
' preestablecida en el sistema será: A123.

Option Explicit

Sub Main()
	Dim passw_control,passw_digitado As String
	passw_control="A123"	' Cadena con contraseña del sistema
	passw_digitado = ""		' Cadena que guardará el password del usuario
	
	'Primer intento de digitar el password
	passw_digitado = InputBox("¿Password?")
	
	' strComp(passw_digitado,passw_control) compara las cadenas de passw_digitado 
	' con passw_control para revisar si son iguales.
	' Esta función de comparación retorna 0 si las cadenas de los 
	' parámetros son iguales. 
	' Por eso en el if se realiza la comparación. Si lo que retorna es igual a 0, 
	' entonces son iguales.
	
	If strComp(passw_digitado,passw_control) = 0 Then
		MsgBox("Acceso Concedido")
	Else
		MsgBox("Error en el password. Lleva 1 intento fallido.")
		' Segundo intento de digitar el password
		passw_digitado = InputBox("¿Password?")
		If strComp(passw_digitado,passw_control) = 0 Then
			MsgBox("Acceso Concedido")
		Else
			MsgBox("Error en el password. Lleva 2 intentos fallido.")
			' Tercer intento de digitar el password
			passw_digitado = InputBox("¿Password?")
			If strComp(passw_digitado,passw_control) = 0 Then
				MsgBox("Acceso Concedido")
			Else
				MsgBox("Ha errado 3 veces. Acceso denegado.")
			End If 	' Cierre del tercer If
		End If 		' Cierre del segundo If
	End If 			' Cierre del primer If
End Sub

