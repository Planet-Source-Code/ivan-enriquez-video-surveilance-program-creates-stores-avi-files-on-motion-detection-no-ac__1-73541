Attribute VB_Name = "mLoguea"
Global ListBuffer As Long 'cuantas entradas guardará la lista
' Sub Loguea. Escribe una línea en modo append en un archivo determinado (lo crea si no existe y agrega si ya existe)
' El archivo puede pasarse (cuando se requiere bitácoras a diferente nivel) o si se omite se usa por default app.path\Log.txt
'
' Modo 0 ó omitido, Agrega fechaHora a la línea. Modo 1 solo escribe el mensaje

Public Sub loguea(cadena As String, Optional modo As Long = 0, Optional archivo As String = "")
Dim canal As Long

On Error GoTo HuboError
canal = FreeFile
If archivo = "" Then archivo = FixPath(App.Path) + App.Title + ".log"
Call Create_Directory(parsefile(archivo, "ruta"))
Open archivo For Append As canal
Select Case modo
   Case 0
      Print #canal, DattedLine(cadena)
   Case 1
      Print #canal, cadena
End Select
Call Log2List(cadena, Form1.List1)
Close #canal
On Error GoTo 0
Exit Sub
' Rutina general de error
HuboError:
MsgBox ("Error at Sub LOGUEA:" + Err.Description + vbCr + vbLf + "Can't open or create Logfile" + vbCr + vbLf + "Archivo:" + archivo + vbCr + vbLf + "This is normally due to lack of access rights on that folder")
End
End Sub


' Usa un listbox para mostrar una bitácora. Para evitar que el programa truene algun dia por memoria,
' maneja la variable global ListBuffer que borra entradas mayores a su tamaño.
'
' Se deben definir los siguientes valores:
' ListBuffer tamaño del buffer. Lineas que guardará en memoria la bitácora

Public Sub Log2List(ByVal cadena As String, ByRef lista As ListBox)
If ListBuffer = 0 Then ListBuffer = 50
lista.AddItem (cadena)
If lista.ListCount > ListBuffer Then
   lista.RemoveItem (0)
End If
lista.ListIndex = lista.ListCount - 1
End Sub

'Agrega Fecha y Hora a una cadena
Public Function DattedLine(ByVal cadena As String) As String
DattedLine = Format(Now, "yyyymmmdd hh:mm:ss") + " " + cadena
End Function

