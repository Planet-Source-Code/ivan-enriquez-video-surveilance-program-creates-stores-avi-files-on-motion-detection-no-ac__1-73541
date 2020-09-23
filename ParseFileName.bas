Attribute VB_Name = "Parsefilename"

Public Function parsefile(ByVal archivo As String, ByVal parte As String) As String
Dim PosSlash, PosPunto, PosColon, LastSlash, I As Long
Dim c As String
PosSlash = 0
LastSlash = 0 'un slash antes del ultimo
PosPunto = 0
PosColon = 0
For I = 1 To Len(archivo)
  c = Mid(archivo, I, 1)
  If c = "." Then PosPunto = I  'Posicion del ULTIMO punto
  If c = "\" Then
     LastSlash = PosSlash
     PosSlash = I  'Posicion del ULTIMO backslash
  End If
  If c = ":" Then PosColon = I  'Posicion del :
Next

'Asumiendo \\mxsfps01\directorio1\directorio2\directorio3\archivo.txt
'y         c:\directorio1\directorio2\directorio3\archivo.txt
Select Case LCase(parte)

'
'c:
Case "drive"
   parsefile = Mid(archivo, 1, PosColon)

'\\mxsfps01\directorio1\directorio2\directorio3\
'c:\directorio1\directorio2\directorio3\
Case "ruta"
   parsefile = Mid(archivo, 1, PosSlash)

'\directorio1\directorio2\directorio3\
'\directorio1\directorio2\directorio3\
Case "rutasindrive"
   parsefile = Mid(archivo, 1, PosSlash)
   If InStr(1, parsefile, ":") Then 'ruta con drive? hay que removerlo
      parsefile = Mid(parsefile, InStr(1, parsefile, ":") + 1)
   Else
     If Mid(parsefile, 1, 2) = "\\" Then 'ruta UNC? remueve el nombre del servidor
        parsefile = Mid(parsefile, 3)
        parsefile = Mid(parsefile, InStr(1, parsefile, "\"))
     End If
   End If
   
'\\mxsfps01\directorio1\directorio2\directorio3\archivo
'c:\directorio1\directorio2\directorio3\archivo
Case "completosinextension"
   parsefile = Mid(archivo, 1, PosPunto - 1)

'archivo
'archivo
Case "nombresinextension"
   parsefile = Mid(archivo, PosSlash + 1, PosPunto - PosSlash - 1)

'archivo.txt
'archivo.txt
Case "nombreconextension"
   parsefile = Mid(archivo, PosSlash + 1)

'txt
'txt
Case "extension"
   parsefile = Mid(archivo, PosPunto + 1)


'Return Parent Directory only (one before the last)
'directorio3
'directorio3
Case "parentdir"
   If LastSlash > 0 Then
      parsefile = Mid(archivo, LastSlash + 1, PosSlash - LastSlash - 1)
   Else
      parsefile = Mid(archivo, 1, PosSlash)
   End If
   

'Return Parent Full Path
'\\mxsfps01\directorio1\directorio2\directorio3
'c:\directorio1\directorio2\directorio3
Case "parentpath"
      parsefile = Mid(archivo, 1, PosSlash - 1)

Case Else
      parsefile = ""
End Select
End Function
