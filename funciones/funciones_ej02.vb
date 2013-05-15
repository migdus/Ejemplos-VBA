' Autor: Miguel Dussán
' Este programa redondea un número hacia abajo y hacia arriba

Option Explicit

Sub main()

    Dim numero As Single
    
    numero = Val(InputBox("Escriba un número"))
    
    ' Arg1 - número real a redondear
    ' Arg 2 - El número de dígitos que se quieren redondear

    MsgBox ("El número " + CStr(numero) + " redondeado hacia abajo: " + CStr(WorksheetFunction.RoundDown(numero, 0)))
    MsgBox ("El número " + CStr(numero) + " redondeado hacia arriba: " + CStr(WorksheetFunction.RoundUp(numero, 0)))
    

End Sub