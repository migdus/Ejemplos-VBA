'Autor: Miguel Dussán

' Este programa suma dos números

Sub Main() ' Inicia función principal

Dim a As Integer ' Declaración de entero (primer número)
Dim b As Integer ' Declaración de entero (segundo número)
Dim c As Integer ' Declaración de entero (almacenar el resultado)

a = 0 ' Inicialización de entero
b = 0 ' Inicialización de entero
c = 0 ' Inicialización de entero

'Solicita el primer número y lo asigna a la variable a
a = InputBox("Ingrese el primer número")

'Solicita el segundo número y lo asigna a la variable b
b = InputBox("Ingrese el segundo número")

c = a + b ' Realiza la operación y la asigna a c

' Variable el resultado de la operación
Dim mensaje_resultado As String

' Las variables numéricas deben convertirse a cadenas
' Para eso se usa la función CStr() que recibe como parámetro
' La variable con el valor numérico, y lo convierte en cadena.
' Las cadenas se concatenan (unen) utilizando el operador "+".

mensaje_resultado = CStr(a) + " + " + CStr(b) + " = " + CStr(c)

'Se muestra el resultado
MsgBox (mensaje_resultado)

End Sub ' Fin de la función principal
