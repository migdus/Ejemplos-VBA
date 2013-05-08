' Autor: Miguel Dussán

' Escriba un programa que capture e imprima una matriz de forma descendente utilizando la matriz original

Option Explicit
Sub main()
    ' Declaración de:
    ' i , j : variables de control para recorrer las posiciones de la matriz
    ' aux: variable auxiliar
    ' filas, columnas: almacela las dimensiones de la matriz según defina el usuario
    ' m: la matriz
    Dim i, j, aux, filas, columnas, m() As Integer
    
    ' band ayuda a controlar el ordenamiento de burbuja
    Dim band As Boolean
    
    ' Solicitar dimensiones de matriz
    filas = Val(InputBox("Número de filas?"))
    
    columnas = Val(InputBox("Número de Columnas?"))
    ' Definir dimensiones de la matriz
    
    ReDim m(filas - 1, columnas - 1)
    
    ' Capturar valores para la matriz
    For i = 0 To filas - 1
        For j = 0 To columnas - 1
            m(i, j) = Val(InputBox("Valor para la posición " + CStr(i) + "," + CStr(j)))
        Next
    Next
    
    ' Realizar el ordenamiento
    
    Do
        ' Se asume que la matriz se encuentra organizada, a menos que se realice
        ' al menos un intercambio de posiciones de los valores
        band = True
        
        ' Iteración por toda la matriz
        For i = 0 To filas - 1
            For j = 0 To columnas - 1
            
                ' Si la columna siguiente a la evaluada es válida/existe
                If j + 1 < columnas Then
                
                    ' Si el valor de la fila, columna siguiente es mayor al valor de la posición actual,
                    ' se hace el intercambio de valores
                    If m(i, j + 1) > m(i, j) Then
                        aux = m(i, j + 1)
                        m(i, j + 1) = m(i, j)
                        m(i, j) = aux
                        band = False ' Falso porque se hizo intercambio de valores
                    End If
                Else
                    ' Este "si-no" se ejecuta solo si la columna siguiente a la evaluada se encuentra
                    ' fuera del rango del tamaño de la matriz
                    
                    ' Si existe una fila posterior
                    If i + 1 < filas Then
                    
                        ' Si el valor de la posición en la siguiente fila, primera columna es mayor al que
                        ' se está evaluando, se realiza el intercambio
                        If m(i + 1, 0) > m(i, j) Then
                            aux = m(i + 1, 0)
                            m(i + 1, 0) = m(i, j)
                            m(i, j) = aux
                            band = False ' Falso porque se hizo intercambio de valores
                        End If
                    End If
                End If
            Next
        Next
    Loop While (band = False)
    
    ' Mostrar resultado en las celdas de la hoja de cálculo
    For i = 0 To filas - 1
        For j = 0 To columnas - 1
            Cells(i + 1, j + 1) = m(i, j)
        Next
    Next

End Sub