' Autor: Miguel Duss�n

' Dise�e un algoritmo que dado el n�mero ingresado por el usuario, muestre en 
' pantalla el d�a de la semana al que corresponde (ejemplo: 3 es mi�rcoles, 
' 7 es domingo). Si el usuario ingresa una opci�n distinta a los n�meros de la 
' semana, deber� mostrar un mensaje de error.
Option Explicit

Sub Main()
	Dim numero_dia As Integer
	' Captura del n�mero de d�a. CInt() convierte la entrada a Integer
	numero_dia = CInt(InputBox("�N�mero del d�a de la semana?"))
	
	' Estructura de selecci�n m�ltiple
	' Se examina el valor de la variable numero_dia
	
	Select Case numero_dia
		Case 1						' Si la variable tiene valor 1 imprime lunes
			MsgBox("Lunes")
		Case 2
			MsgBox("Martes")
		Case 3
			MsgBox("Mi�rcoles")
		Case 4
			MsgBox("Jueves")
		Case 5
			MsgBox("Viernes")
		Case 6
			MsgBox("S�bado")
		Case 7
			MsgBox("Domingo")		
		Case Else ' Si no coincide ninguno de los valores con el contenido de la variable
			MsgBox("No existe un d�a asociado a ese n�mero")		
	End Select
	
End Sub
#include <iostream>
	
using namespace std;

int main(){

	

	switch(numero_dia){
		case 1: 										//Si la variable tiene valor 1
			cout << "Lunes" << endl;	//Imprime "Lunes"
			break;										//Termina el bloque que se ejecuta para el caso 1
		case 2: 										//Si la variable tiene valor 2
			cout << "Martes"<< endl;  //Imprime "Martes"
			break;										//Termina el bloque que se ejecuta para el caso 2
		case 3: 										//Si la variable tiene valor 3
		cout << "Miercoles"<< endl;	//Imprime "Mi�rcoles"
		break;										//Termina el bloque que se ejecuta para el caso 3
		case 4: 										//Si la variable tiene valor 4
		cout << "Jueves"<< endl;	//Imprime "Jueves"
		break;										//Termina el bloque que se ejecuta para el caso 4
		case 5: 										//Si la variable tiene valor 5
		cout << "Viernes"<< endl;	//Imprime "Viernes"
		break;										//Termina el bloque que se ejecuta para el caso 5
		case 6: 										//Si la variable tiene valor 6
		cout << "Sabado"<< endl;	//Imprime "Sabado"
		break;										//Termina el bloque que se ejecuta para el caso 6
		case 7: 										//Si la variable tiene valor 7
		cout << "Domingo"<< endl;	//Imprime "Domingo"
		break;										//Termina el bloque que se ejecuta para el caso 7
		default: 										//Si  el usuario escribe otro n�mero
		cout << "Error"<< endl;		//Imprime "Error"
		break;										//termina el bloque que se ejecuta si no coinciden las opciones anteriores
	}
	

	return 0;
}
