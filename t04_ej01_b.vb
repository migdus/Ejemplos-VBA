'Autor: Miguel Dussán

' Este programa determina si un número entero es negativo
' utilizando una estructura Sí-Entonces- Si no

Option Explicit ' Para forzar declaración de variables

Sub Main()

	Dim numero As Integer ' Declaración de entero
	Dim cadena_temp As String ' Declaración de cadena temporal, se utiliza para
														' guardar los datos capturados por teclado, y luego
														' convertirlos al tipo de dato Integer

	numero = 0 ' Inicialización
	cadena_temp = "" ' Inicialización de cadena con una cadena vacía

	'La línea a continuación muestra en pantalla la cadena que se
	'encuentra dentro de las comillas dobles, utilizando el comando
	' InputBox(). Además muestra un campo para que el usuario ingrese
	' un valor que será asignado a la variable numero.
	' La salida de InputBox se asigna a una cadena, y luego esta se convierte
	' a entero con la función CInt 

	cadena_temp = InputBox("Escriba un número")
	numero = CInt(cadena_temp)


	' Estructura selectiva Sí-Entonces- Si no (If-Then-Else).
	' El alcance de la estructura se identifica cierra con las
	' palabras clave: "End If"
	' La condición para que ejecute el código dentro de la estructura es
	' que el número sea menor que cero.

	If numero < 0 Then					' Abre la estructura
		MsgBox("El número es negativo")	' Imprime un mensaje si el número es negativo
	Else								' Abre un Si No
		MsgBox("El número es positivo")	' Imprime un mensaje si el número es positivo	
	End If								' Cierra la estructura

End Sub