' Autor: Miguel Duss�n

' Escriba un programa que determine si un n�mero entero n es impar

Option Explicit ' Para forzar declaraci�n de variables

Sub Main()
	Dim numero As Integer	' Declaraci�n de entero
	numero = 0				' Inicializaci�n
	
	' Impresi�n y captura de entrada en la variable numero
	numero = InputBox("Escriba un n�mero") 
	
	' Estructura selectiva con la condici�n:
	' si el n�mero MOD 2 es diferente de cero, entonces el n�mero es impar
	If numero mod 2 <> 0 Then
		MsgBox("El n�mero es impar")
	End If
End Sub


