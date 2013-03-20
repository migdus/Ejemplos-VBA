' Autor: Miguel Duss�n

' Dise�e un programa que calcule e imprima en pantalla el di�metro,
' la longitud y el �rea de un c�rculo, dado r ingresado por el usuario.
' El programa debe asumir que PI = 3.14159.

Option Explicit ' Para forzar declaraci�n de variables

Sub Main()
    ' Declaraci�n de variables como n�meros de coma flotante
    ' (n�meros decimales)
    Dim pi, diametro, longitud, area As Single
    Dim r As Integer ' Declaraci�n de entero
		
    pi = 3.14159    ' Inicializaci�n de Pi
    diametro = 0    ' Inicializaci�n
    longitud = 0    ' Inicializaci�n
    area = 0        ' Inicializaci�n
		
		'Los pasos a continuaci�n se pueden poner en uno solo, y as�
		'no utilizar cadena_temp como en los ejercicios anteriores
		'Dim cadena_temp As String
		'cadena_temp = ""
		'cadena_temp  = InputBox("Ingrese el radio del c�rculo")
    'r = CInt(cadena_temp)
		'En la siguiente l�nea se pone la funci�n que muestra la ventana que pide
		'que se introduzca algo (InputBox) dentro de la funci�n que convierte String
		'a Integer. As�, lo que escriba el usuario y retorne InputBox se convierte en
		'la entrada de la funci�n CInt
	
		r = CInt(InputBox("Ingrese el radio del c�rculo"))
    
    ' C�lculos
    diametro = 2 * r
    longitud = pi * 2 * r
    area = pi * r ^ 2 ' El s�mbolo ^ es para hallar potencias
    
    Dim resultado As String ' Variable para almacenar los resultados
    
	' La variable para almacenar los resultados va a "acumular" los valores para
	' todos los c�lculos. Esto se hace uniendo diferentes trozos del texto que
	' queremos que aparezca en la ventana que muestra el resultado.
	' Para concatenar (unir) cadenas se utiliza el operador "+".
	
	' La funci�n CStr() se utiliza para volver el contenido de una variable
	' num�rica en una cadena de caracteres. Toma como par�metro la variable
	' num�rica.
	
	' vbTab es un caracter especial que se interpreta como si se hiciera "tab"
	' con el teclado. Sirve para alinear cadenas. Aqu� se usa para que los
	' n�meros se muestren alineados.
	' Mnemotecnia: "visual basic" vb "Tab" Tab
	
	' vbCrLf es un caracter especial que se interpreta como un fin de l�nea.
	' Si este no se utilizara, el resultado se mostrar�a en una sola l�nea
	' Mnemotecnia: "visual basic" vb "Carriage return Line feed" CrLf
	
    resultado = "El di�metro es:" + vbTab + CStr(diametro) + vbCrLf
    resultado = resultado + "La longitud es:" + vbTab + CStr(longitud) + vbCrLf
    resultado = resultado + "El �rea es:" + vbTab + vbTab + CStr(area) + vbCrLf
    
    MsgBox (resultado)
End Sub
