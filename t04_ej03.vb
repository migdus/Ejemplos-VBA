' Autor: Miguel Dussán

' Escriba un programa que determine si un número es positivo y menor que 100

Option Explicit ' Para forzar declaración de variables

Sub Main()
	Dim numero as Integer	' Declaración de variable
	Dim cadena_temp As String 
	
	numero = 0				' Inicialización de variable
	cadena_temp = ""
	
	cadena_temp = InputBox("Escriba un número") 'Solicitud de un número
	numero = CInt(cadena_temp)
	
	' Estructura selectiva con una condicional compuesta
	' Tiene dos condiciones unidas por un Y lógico
	
	If numero > 0 And numero < 100 Then
		MsgBox("El número es positivo y menor que 100")
	End If
End Sub


