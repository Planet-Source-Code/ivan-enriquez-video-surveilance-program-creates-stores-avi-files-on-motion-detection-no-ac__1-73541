Attribute VB_Name = "mFixPath"


' Recibe un path y lo regresa con el ultimo "\". Algunas funciones de vb regresan o no ese
' ultimo backslash, y dependiendo de lo que regresen tambien varia si lo trae o no. Por
' ejemplo, dir1.path solo regresa el backslash si es una unidad raiz, y el resto no la
' incluye, y tambien en unidades de red cambia la forma como funciona.
' esta funcion elimina todos esos problemas agregandole si necesita ese ultimo backslash
' Ivan Enriquez Muñoz 24 nov 2005


Function FixPath(ByVal ruta As String) As String
If right(ruta, 1) = "\" Then
   FixPath = ruta
Else
   FixPath = ruta + "\"
End If
End Function
