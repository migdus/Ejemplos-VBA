' Autor: Miguel Dussán
' Este programa implementa una función que computa la adición de dos números.

Option Explicit

Sub main()

    Dim s1, s2, s3 As Integer
    
    s1 = capturaNumero()
    
    s2 = capturaNumero()
    
    Call suma(s1, s2, s3)

    MsgBox (CStr(s1) + " + " + CStr(s2) + " = " + CStr(s3))

End Sub


Sub suma(ByVal n1 As Integer, ByVal n2 As Integer, ByRef n3 As Integer)

    n3 = n1 + n2

End Sub

Function capturaNumero() As Integer

    capturaNumero = Val(InputBox("Escriba un número"))
    
End Function

