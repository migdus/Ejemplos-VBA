' Autor: Miguel Dussán

' Este programa determina el número que más se repite dentro de los
' ingresados por un usuario en un arreglo.

Option Explicit

Sub main()
    Dim a() As Integer ' Arreglo de valores ingresados
    Dim frecuencias() As Integer ' Arreglo de frecuencias
    Dim tam As Integer ' Tamaño del arreglo
    Dim num_mayor As Integer ' Número más grande de los ingresados
    Dim frec_alta As Integer ' Frecuencia más alta
    Dim i As Integer ' Variable de control de ciclo
    
    tam = 0
    num_mayor = -9999 ' Un valor hipotéticamente pequeño
    frec_alta = -9999
    
    tam = Val(InputBox("Cuantos numeros desea ingresar?"))
    ReDim a(tam)
    
    'Captura de los valores
    For i = 1 To tam
    ' ciclo do-while para evitar que el usuario ingrese números negativos'
        Do
            a(i) = Val(InputBox("Ingrese el numero " + CStr(i) + " (positivos solamente): "))
        Loop While a(i) < 0
        ' Encontrar si el número ingresado es el más grande
        If a(i) > num_mayor Then
            num_mayor = a(i)
        End If
    Next
    
    ' Definición de dimensión del arreglo con las frecuencias
    ' Se debe hacer en este momento, porque antes no se tenía el tamaño del mismo.
    ' El índice de este arreglo representará al número, y el valor la frecuencia.
    ' Por esto se declara con el número: num_mayor+1
    ReDim frecuencias(num_mayor + 1)
    
    ' Inicialización del arreglo de frecuencias en ceros
    For i = 0 To num_mayor
        frecuencias(i) = 0
    Next
    
    ' Calcular las frecuencias para todos los números
    
    For i = 0 To tam
    ' El índice utilizado es el valor ingresado por el usuario.
    ' El valor es el incremento, porque quiere decir que se ha encontrado
    ' una coincidencia.
        frecuencias(a(i)) = frecuencias(a(i)) + 1
    Next
    
    ' Revisar cuál es la frecuencia más alta
    For i = 0 To num_mayor
    If frecuencias(i) > frec_alta Then
    frec_alta = i
    End If
    Next
    
    MsgBox ("El numero con la frecuencia mas alta fue: " + CStr(frec_alta) + " con una frecuencia de: " + CStr(frecuencias(frec_alta)))
End Sub

