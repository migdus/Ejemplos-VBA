'Autor: Miguel Duss�n

' Este programa suma dos n�meros

Sub Main() ' Inicia funci�n principal

Dim a As Integer ' Declaraci�n de entero (primer n�mero)
Dim b As Integer ' Declaraci�n de entero (segundo n�mero)
Dim c As Integer ' Declaraci�n de entero (almacenar el resultado)

a = 0 ' Inicializaci�n de entero
b = 0 ' Inicializaci�n de entero
c = 0 ' Inicializaci�n de entero

'Solicita el primer n�mero y lo asigna a la variable a
a = InputBox("Ingrese el primer n�mero")

'Solicita el segundo n�mero y lo asigna a la variable b
b = InputBox("Ingrese el segundo n�mero")

c = a + b ' Realiza la operaci�n y la asigna a c

' Variable el resultado de la operaci�n
Dim mensaje_resultado As String

' Las variables num�ricas deben convertirse a cadenas
' Para eso se usa la funci�n CStr() que recibe como par�metro
' La variable con el valor num�rico, y lo convierte en cadena.
' Las cadenas se concatenan (unen) utilizando el operador "+".

mensaje_resultado = CStr(a) + " + " + CStr(b) + " = " + CStr(c)

'Se muestra el resultado
MsgBox (mensaje_resultado)

End Sub ' Fin de la funci�n principal
