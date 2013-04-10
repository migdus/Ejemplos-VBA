' Autor: Miguel Dussán

Option Explicit

Sub main()
	Dim a() As Integer ' Declaración del arreglo de tipo entero, aun no se conoce el tamaño
	Dim tam As Integer ' Tamaño del arreglo

	Dim promedio As Variant ' Variable acumuladora
	Dim i As Integer ' Variable de control de ciclo

	tam = Val(InputBox("Cuantos numeros desea ingresar?"))

	ReDim a(tam) 'Como ahora se conoce el tamaño, se fija con el comando ReDim

	' Captura de los valores
	For i = 1 To tam
	    a(i) = Val(InputBox("Ingrese el número " + CStr(i)))
	Next

	' Mostrar el valor de la información almacenada en el arreglo
	Cells(1, 1) = "Posición"
	Cells(1, 2) = "Valor"

	For i = 0 To tam
		' Impresión del índice del arreglo y del valor (j y a(j), respectivamente)
		Cells(i + 2, 1) = i
		Cells(i + 2, 2) = a(i)

		' Acumulador para calcular el promedio.
		' Se aprovecha la iteración en la que se muestran los numeros para acumular
		' los valores de cada una de las posiciones del arreglo en una variable.

		promedio = promedio + a(i)
	Next

	Cells(tam + 2, 1) = "El promedio de los numeros fue: " + CStr(promedio)
End Sub
