' Autor: Miguel Dussán

'Escriba un programa que permita al usuario introducir N números enteros, y 
' permita determinar el mayor y el menor de los números introducidos.

Option Explicit

Sub Main()
	' cantidad_numeros captura cuántos números quiere introducir el usuario
	' numero captura un número introducido por el usuario
	' num_mayor almacena el número más grande detectado
	' num_menor captura el número más pequeño detectado
	' cont es la variable de control de ciclo
	Dim cantidad_numeros, numero, numero_mayor, numero_menor, cont As Integer
	
	'Inicialización
	cantidad_numeros = 0
	numero = 0
	cont = 0
	
	'Inicialización del número mayor y del número menor
	'Se inicializan con los valores más grande y más pequeño para
	'ese tipo de datos
	
	' Constante que indica el número más grande de
	' los enteros. Se asigna este valor para que cualquier 
	'número lo supere y sea, en principio, el mayor
	numero_mayor = -32768
	
	' Constante que indica el número más pequeño de
	' los enteros. Se asigna este valor para que cualquier 
	'número esté por debajo y sea, en principio, el menor
	numero_menor = 32767
	
	' Pedir la cantidad de números
	cantidad_numeros = InputBox("¿Cuántos números desea ingresar?")
	
	' Mientras el contador sea menor a la cantidad de números
	While cont < cantidad_numeros
		numero = InputBox("¿Número?")
		
		' Revisar si este número es mayor al número más grande guardado
		If numero > numero_mayor Then
			' Si es mayor, se guarda como el mayor de los ingresados
			numero_mayor = numero
		End If
		' Revisar si este número es más pequeño que el menor guardado
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
	
	' Mostrar los números en pantalla
	MsgBox(resultado)
End Sub
