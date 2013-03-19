'Autor: Miguel Duss�n

' Este programa determina si un n�mero entero es negativo
' utilizando una estructura S�-Entonces

Sub Main() ' Inicio del proceso de nombre Main

	Dim numero As Integer ' Declaraci�n de entero
	numero = 0 ' Inicializaci�n

	'La l�nea a continuaci�n muestra en pantalla la cadena que se
	'encuentra dentro de las comillas dobles, utilizando el comando
	' InputBox(). Adem�s muestra un campo para que el usuario ingrese
	' un valor que ser� asignado a la variable numero. 

	numero = InputBox("Escriba un n�mero")

	' Estructura selectiva S�-Entonces (If-Then).
	' El alcance de la estructura se identifica cierra con las
	' palabras clave: "End If"
	' La condici�n para que ejecute el c�digo dentro de la estructura es
	' que el n�mero sea menor que cero.

	If numero < 0 Then					' Abre la estructura
		MsgBox("El n�mero es negativo")	' Imprime un mensaje si el n�mero es negativo
	End If								' Cierra la estructura

End Sub ' Fin del proceso