'Autor: Miguel Duss�n

' Dise�e un programa que emule un control de acceso. El programa debe preguntar al
' usuario por una contrase�a alfanum�rica (es decir, que puede contener n�meros y 
' letras) y verificar si la contrase�a es correcta. El usuario dispondr� de tres 
' intentos para acertar la contrase�a. Si el usuario escribe la contrase�a de 
' forma correcta, el programa debe mostrar un mensaje que diga: Acceso concedido; 
' de lo contrario mostrar� el mensaje: Acceso denegado. La contrase�a que estar� 
' preestablecida en el sistema ser�: A123.

Option Explicit

Sub Main()
	Dim passw_control,passw_digitado As String
	passw_control="A123"	' Cadena con contrase�a del sistema
	passw_digitado = ""		' Cadena que guardar� el password del usuario
	
	'Primer intento de digitar el password
	passw_digitado = InputBox("�Password?")
	
	' strComp(passw_digitado,passw_control) compara las cadenas de passw_digitado 
	' con passw_control para revisar si son iguales.
	' Esta funci�n de comparaci�n retorna 0 si las cadenas de los 
	' par�metros son iguales. 
	' Por eso en el if se realiza la comparaci�n. Si lo que retorna es igual a 0, 
	' entonces son iguales.
	
	If strComp(passw_digitado,passw_control) = 0 Then
		MsgBox("Acceso Concedido")
	Else
		MsgBox("Error en el password. Lleva 1 intento fallido.")
		' Segundo intento de digitar el password
		passw_digitado = InputBox("�Password?")
		If strComp(passw_digitado,passw_control) = 0 Then
			MsgBox("Acceso Concedido")
		Else
			MsgBox("Error en el password. Lleva 2 intentos fallido.")
			' Tercer intento de digitar el password
			passw_digitado = InputBox("�Password?")
			If strComp(passw_digitado,passw_control) = 0 Then
				MsgBox("Acceso Concedido")
			Else
				MsgBox("Ha errado 3 veces. Acceso denegado.")
			End If 	' Cierre del tercer If
		End If 		' Cierre del segundo If
	End If 			' Cierre del primer If
End Sub

