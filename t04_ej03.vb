' Autor: Miguel Duss�n

' Escriba un programa que determine si un n�mero es positivo y menor que 100

Option Explicit ' Para forzar declaraci�n de variables

Sub Main()
	Dim numero as Integer	' Declaraci�n de variable
	numero = 0				' Inicializaci�n de variable
	
	numero = InputBox("Escriba un n�mero") 'Solicitud de un n�mero
	
	' Estructura selectiva con una condicional compuesta
	' Tiene dos condiciones unidas por un Y l�gico
	
	If numero > 0 And numero < 100 Then
		MsgBox("El n�mero es positivo y menor que 100")
	End If
End Sub


