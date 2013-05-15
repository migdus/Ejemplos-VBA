' Autor: Miguel Dussán
' Este programa implementa una función que computa la adición de dos números.

Option Explicit

Sub main()

    Dim s1, s2, s3 As Integer
    
    s1 = capturaNumero()
    
    s2 = capturaNumero()
    
    s3 = suma(s1, s2)

    MsgBox (CStr(s1) + " + " + CStr(s2) + " = " + CStr(s3))

End Sub

Function suma(ByVal n1 As Integer, ByVal n2 As Integer) As Integer

    suma = n1 + n2

End Function

Function capturaNumero() As Integer

    capturaNumero = Val(InputBox("Escriba un número"))
    
End Function

