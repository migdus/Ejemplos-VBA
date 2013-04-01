Option Explicit

Sub Main()
Dim j1, j2 As Integer

j1 = CInt(InputBox("Jugada para jugador 1? 1. Piedra 2. Papel 3. Tijeras"))

j2 = CInt(InputBox("Jugada para jugador 2? 1. Piedra 2. Papel 3. Tijeras"))

If j1 = 1 Then
    If j2 = 1 Then
        MsgBox ("Empate")
    Else
        If j2 = 2 Then
            MsgBox ("Gana J2")
        Else
            If j2 = 3 Then
                MsgBox ("Gana J1")
            End If
        End If
    End If
End If

If j1 = 2 Then
    If j2 = 1 Then
        MsgBox ("Gana J1")
    Else
        If j2 = 2 Then
            MsgBox ("Empate")
        Else
            If j2 = 3 Then
                MsgBox ("Gana J3")
            End If
        End If
    End If
End If

If j1 = 3 Then
    If j2 = 1 Then
        MsgBox ("Gana J2")
    Else
        If j2 = 2 Then
            MsgBox ("Gana J1")
        Else
            If j2 = 3 Then
                MsgBox ("Empate")
            End If
        End If
    End If
End If

End Sub