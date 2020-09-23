Attribute VB_Name = "Commonlib"

Function lista(ByVal lCadena As String, ByVal lElemento As Long, ByVal lSeparador As String) As Variant
Dim LocalCadena, LocalSeparador As String
Dim LocalElementos, PosicionInicial, PosicionFinal As Long
LocalCadena = lCadena: LocalSeparador = "°"
LocalCadena = Replace(LocalCadena, lSeparador, LocalSeparador)
If Left$(LocalCadena, 1) <> LocalSeparador Then LocalCadena = LocalSeparador + LocalCadena
If Right(LocalCadena, 1) <> LocalSeparador Then LocalCadena = LocalCadena + LocalSeparador
LocalElementos = 0: PosicionInicial = 0: PosicionFinal = 0
For i = 1 To Len(LocalCadena)
   If Mid$(LocalCadena, i, 1) = LocalSeparador Then
      LocalElementos = LocalElementos + 1
      If lElemento = LocalElementos Then PosicionInicial = i
      If lElemento + 1 = LocalElementos Then PosicionFinal = i
   End If
Next i
If lElemento = 0 Then
   lista = LocalElementos - 1
Else
   If lElemento > LocalElementos - 1 Then
      lista = "ErrorOverflow"
   Else
      lista = Mid$(LocalCadena, PosicionInicial + 1, PosicionFinal - PosicionInicial - 1)
   End If
End If
End Function

Public Function Maximo(ByVal valor1 As Double, ByVal valor2 As Double) As Double
 Maximo = valor2
 If valor1 > valor2 Then Maximo = valor1
End Function

Public Function Minimo(ByVal valor1 As Long, ByVal valor2 As Long) As Long
 Minimo = valor2
 If valor1 < valor2 Then Minimo = valor1
End Function

Public Function LastPos(cadena As String, caracter As String) As Long
Dim i, posicion As Long
posicion = 0
For i = 1 To Len(cadena)
   If Mid$(cadena, i, 1) = caracter Then posicion = i
Next i
LastPos = posicion
End Function

Public Sub Progress(ByVal P1 As PictureBox, ByVal Nvalor As Long, ByVal MaxVal As Long, Optional Barcolor As ColorConstants, Optional fondo As ColorConstants)
Dim valor As Long
If Barcolor = 0 Then
   P1.FillColor = vbRed
Else
   P1.FillColor = Barcolor
End If
P1.FillStyle = 0
If fondo > 0 Then P1.BackColor = fondo

'P1.ForeColor = Barcolor
If Nvalor >= 0 And Nvalor <= MaxVal Then
   valor = Nvalor * P1.Width / MaxVal
   P1.Cls
   P1.Line (0, 0)-(valor, P1.Top + P1.Height), Barcolor, B 'HOR
End If
End Sub

