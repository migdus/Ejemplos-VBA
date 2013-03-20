' Autor: Miguel Duss�n

' Imprima todos los n�meros impares entre el n�mero 1 y el 200. De forma intercalada,
' uno del principio y uno del final. Ejemplo: 1,199,3,197,5,195,...

Option Explicit

Sub Main()
    Dim cont_asc, cont_desc As Integer
    Dim resultado As String
    resultado = ""      ' Almacena todo el resultado para mostrarlo al final
                        ' Funciona como un acumulador de cadenas
    cont_asc = 1    ' contador ascendente
    cont_desc = 200 ' contador descendente
    
    ' Ciclo while. La condici�n es que el contador ascendente sea
    ' Menor o igual a 200. Tambi�n se puede colocar que el contador descendente
    ' sea mayor o igual a 1, o ambas.
    
    While cont_asc <= 200
        ' Condici�n para el contador ascendente.
        ' Si el n�mero ascendente es impar.
        ' Se usa la funci�n CStr() para convertir el valor de la variable a
        ' cadena y as� poderla concatenar
        If cont_asc Mod 2 <> 0 Then
            resultado = resultado + CStr(cont_asc) + " "
        End If
        
        ' Condici�n para el contador descendente.
        ' Si el n�mero descendente es impar.
        ' Se usa la funci�n CStr() para convertir el valor de la variable a
        ' cadena y as� poderla concatenar
        If cont_desc Mod 2 <> 0 Then
            resultado = resultado + CStr(cont_desc)+ " "
        End If
        
        cont_asc = cont_asc + 1     'Incremento del contador ascendente
        cont_desc = cont_desc - 1   ' Decremento del contador descendente
        
    Wend
    
    MsgBox (resultado)
End Sub

