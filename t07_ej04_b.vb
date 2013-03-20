' Autor: Miguel Dussán

' Escriba un programa que utilizando una sola estructura repetitiva, imprima
' todos los números entre un rango determinado por el usuario, incluyendo esos dos
' números, con incrementos de n. El usuario puede ingresar los rangos inferior y
' superior en cualquier orden.

Option Explicit

Sub Main()
    Dim rango_inferior, rango_superior, incremento, aux As Integer
    
    rango_inferior = 0
    rango_superior = 0
    incremento = 0
    aux = 0

    rango_inferior = CInt(InputBox("¿Primer número del rango?"))
    
    rango_superior = CInt(InputBox("¿Segundo número del rango?"))
    
    incremento = CInt(InputBox("¿Incremento?"))
    
    If rango_superior < rango_inferior Then
        aux = rango_inferior
        rango_inferior = rango_superior
        rango_superior = aux
    End If
    
    Dim resultado As String
    resultado = ""
    
    While rango_inferior <= rango_superior
            resultado = resultado + CStr(rango_inferior) + " "
            rango_inferior = rango_inferior + incremento
    Wend
    
    MsgBox (resultado)
    
End Sub
