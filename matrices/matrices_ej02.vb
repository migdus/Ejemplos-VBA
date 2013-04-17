' Autor: Miguel Dussán

' Escriba un programa que permita crear una matriz del tamaño especificado por el usuario,
' y le permita introducir valores numéricos en ella.

Option Explicit

Sub main()
    Dim m() As Integer ' Matriz
    
    Dim num_filas, num_columnas As Integer ' Almacenan las dimensiones de la matriz
    
    Dim i, j As Integer ' Variables de control del ciclo
    
    num_filas = Val(InputBox("¿Cuántas filas desea introducir?"))
    
    num_columnas = Val(InputBox("¿Cuántas columnas desea introducir?"))
    
    'Definición de matriz
    'Se utilizan los valores del número de filas y de columnas capturados anteriormente para
    'declarar la matriz de forma dinámica.
    'Al valor de número de filas y columnas se les resta uno, pues los índices arrancan en cero.
    
    ReDim m(0 To num_filas - 1, 0 To num_columnas - 1)

    ' Ciclo para capturar los valores de la matriz
    For i = 0 To num_filas - 1 ' Controla la posición de la fila
        For j = 0 To num_columnas - 1 ' Controla la posición de la columna
            m(i, j) = Val(InputBox("Escriba un valor para la posición " + CStr(i) + "," + CStr(j)))
        Next
    Next
    
    ' Mostrar los valores en la hoja de cálculo
    ' Primer ciclo: recorre la matriz por filas
    For i = 0 To num_filas - 1
        ' Segundo ciclo: recorre la matriz por columnas
        For j = 0 To num_columnas - 1
            ' Escribir valores en la hoja de cálculo
            Cells(i + 1, j + 1) = m(i, j)
        Next
    Next

End Sub
