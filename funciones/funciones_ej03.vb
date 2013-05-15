' Autor: Miguel Dussán
' Este programa simula el lanzamiento de un dado

Option Explicit

Sub main()

    Dim resultado As Single
    
    resultado = Rnd() Mod 6 + 1

    MsgBox ("El resultado del dado fue: " + CStr(resultado))

End Sub