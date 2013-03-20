' Autor: Miguel Duss�n

' Escriba un programa que determine si un n�mero entero n es impar

Option Explicit ' Para forzar declaraci�n de variables

Sub Main()
	Dim numero As Integer	' Declaraci�n de entero
	Dim cadena_temp As String ' Cadena temporal
	numero = 0				' Inicializaci�n
	
	' Impresi�n y captura de entrada en la variable numero
	cadena_temp = InputBox("Escriba un n�mero") 
	numero = CInt(cadena_temp)
	
	' Estructura selectiva con la condici�n:
	' si el n�mero MOD 2 es diferente de cero, entonces el n�mero es impar
	If numero mod 2 <> 0 Then
		MsgBox("El n�mero es impar")
	End If
End Sub


