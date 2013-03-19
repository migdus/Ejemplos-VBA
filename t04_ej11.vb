' Autor: Miguel Dussán

' Diseñe un programa que calcule e imprima en pantalla el diámetro,
' la longitud y el área de un círculo, dado r ingresado por el usuario.
' El programa debe asumir que PI = 3.14159.

Sub Main()
    ' Declaración de variables como números de coma flotante
    ' (números decimales)
    Dim pi, diametro, longitud, area As Single
    Dim r As Integer ' Declaración de entero
    pi = 3.14159    ' Inicialización de Pi
    diametro = 0    ' Inicialización
    longitud = 0    ' Inicialización
    area = 0        ' Inicialización
    
    r = InputBox("Ingrese el radio del círculo")
    
    ' Cálculos
    diametro = 2 * r
    longitud = pi * 2 * r
    area = pi * r ^ 2 ' El símbolo ^ es para hallar potencias
    
    Dim resultado As String ' Variable para almacenar los resultados
    
	' La variable para almacenar los resultados va a "acumular" los valores para
	' todos los cálculos. Esto se hace uniendo diferentes trozos del texto que
	' queremos que aparezca en la ventana que muestra el resultado.
	' Para concatenar (unir) cadenas se utiliza el operador "+".
	
	' La función CStr() se utiliza para volver el contenido de una variable
	' numérica en una cadena de caracteres. Toma como parámetro la variable
	' numérica.
	
	' vbTab es un caracter especial que se interpreta como si se hiciera "tab"
	' con el teclado. Sirve para alinear cadenas. Aquí se usa para que los
	' números se muestren alineados.
	' Mnemotecnia: "visual basic" vb "Tab" Tab
	
	' vbCrLf es un caracter especial que se interpreta como un fin de línea.
	' Si este no se utilizara, el resultado se mostraría en una sola línea
	' Mnemotecnia: "visual basic" vb "Carriage return Line feed" CrLf
	
    resultado = "El diámetro es:" + vbTab + CStr(diametro) + vbCrLf
    resultado = resultado + "La longitud es:" + vbTab + CStr(longitud) + vbCrLf
    resultado = resultado + "El área es:" + vbTab + vbTab + CStr(area) + vbCrLf
    
    MsgBox (resultado)
End Sub
