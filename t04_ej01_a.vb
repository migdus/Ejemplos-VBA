'Autor: Miguel Duss�n

' Este programa determina si un n�mero entero es negativo
' utilizando una estructura S�-Entonces

Option Explicit ' Para forzar declaraci�n de variables

Sub Main() ' Inicio del proceso de nombre Main

	Dim numero As Integer ' Declaraci�n de entero
	Dim cadena_temp As String ' Declaraci�n de cadena temporal, se utiliza para
														' guardar los datos capturados por teclado, y luego
														' convertirlos al tipo de dato Integer
	
	numero = 0 ' Inicializaci�n
	cadena_temp = "" ' Inicializaci�n de cadena con una cadena vac�a

	'La l�nea a continuaci�n muestra en pantalla la cadena que se
	'encuentra dentro de las comillas dobles, utilizando el comando
	' InputBox(). Adem�s muestra un campo para que el usuario ingrese
	' un valor que ser� asignado a la variable numero.
	' La salida de InputBox se asigna a una cadena, y luego esta se convierte
	' a entero con la funci�n CInt 

	cadena_temp = InputBox("Escriba un n�mero")
	numero = CInt(cadena_temp)

	' Estructura selectiva S�-Entonces (If-Then).
	' El alcance de la estructura se identifica cierra con las
	' palabras clave: "End If"
	' La condici�n para que ejecute el c�digo dentro de la estructura es
	' que el n�mero sea menor que cero.

	If numero < 0 Then					' Abre la estructura
		MsgBox("El n�mero es negativo")	' Imprime un mensaje si el n�mero es negativo
	End If								' Cierra la estructura

End Sub ' Fin del proceso