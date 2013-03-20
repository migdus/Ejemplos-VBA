' Autor: Miguel Duss�n

'Escriba un programa que permita al usuario introducir N n�meros enteros, y 
' permita determinar el mayor y el menor de los n�meros introducidos.

Option Explicit

Sub Main()
	' cantidad_numeros captura cu�ntos n�meros quiere introducir el usuario
	' numero captura un n�mero introducido por el usuario
	' num_mayor almacena el n�mero m�s grande detectado
	' num_menor captura el n�mero m�s peque�o detectado
	' cont es la variable de control de ciclo
	Dim cantidad_numeros, numero, numero_mayor, numero_menor, cont As Integer
	
	'Inicializaci�n
	cantidad_numeros = 0
	numero = 0
	cont = 0
	
	'Inicializaci�n del n�mero mayor y del n�mero menor
	'Se inicializan con los valores m�s grande y m�s peque�o para
	'ese tipo de datos
	
	' Constante que indica el n�mero m�s grande de
	' los enteros. Se asigna este valor para que cualquier 
	'n�mero lo supere y sea, en principio, el mayor
	numero_mayor = -32768
	
	' Constante que indica el n�mero m�s peque�o de
	' los enteros. Se asigna este valor para que cualquier 
	'n�mero est� por debajo y sea, en principio, el menor
	numero_menor = 32767
	
	' Pedir la cantidad de n�meros
	cantidad_numeros = CInt(InputBox("�Cu�ntos n�meros desea ingresar?"))
	
	' Mientras el contador sea menor a la cantidad de n�meros
	While cont < cantidad_numeros
		numero = CInt(InputBox("�N�mero?"))
		
		' Revisar si este n�mero es mayor al n�mero m�s grande guardado
		If numero > numero_mayor Then
			' Si es mayor, se guarda como el mayor de los ingresados
			numero_mayor = numero
		End If
		' Revisar si este n�mero es m�s peque�o que el menor guardado
		If numero < numero_menor Then
			' Si es menor, se guarda como el menor de los ingresados
			numero_menor = numero
		End If
		
		cont = cont + 1 ' Incrementa en 1 el contador.
	WEnd

	' Preparar los resultados para mostrarlos en pantalla
	
	Dim resultado As String
	
	resultado =  "El numero mayor entre los ingresados es: " + CStr(numero_mayor) + vbCrlf
	resultado =  resultado + "El numero menor entre los ingresados es: " + CStr(numero_menor)
	
	' Mostrar los n�meros en pantalla
	MsgBox(resultado)
End Sub
