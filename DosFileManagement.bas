Attribute VB_Name = "dOSfILEmANAGEMENT"
Option Explicit
'FUNCIONES DE USO COMUN PARA MANEJO DE ARCHIVOS
'Function APItoString(s As String) As String  quita los chr(0) a la der. de una string API
'Function FileSize(sPath As String) As Double
'Function DirSpace(sPath As String) As Double
'Function HextoDec(ByVal cadena As String) As Double
'Function DecToHex(ByVal numero As Long) As String
'Function DirExists(ByVal DirName As String) As Boolean
'Function FileExists(ByVal archivo As String) As Boolean
'Public Sub BorraArchivo(ByVal archivo As String)
'Public Sub copiarchivo(ByVal origen As String, ByVal destino As String)

'API constants
Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

'API types
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'API function calls
Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long



'Truncate a string returned by API calls to the first null char Chr(0)
Function APItoString(s As String) As String
    Dim X As Integer

    X = InStr(s, Chr(0))
    If X <> 0 Then
        APItoString = Left(s, X - 1)
    Else
        APItoString = s
    End If
End Function

Function FileSize(ByVal sPath As String) As Double
    Dim f As WIN32_FIND_DATA
    Dim hFile As Long
    Dim hSize As Long
    FileSize = 0
    'Add the slash to the search path
    'sPath = FixPath(sPath)
    hFile = FindFirstFile(sPath, f)
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    If (f.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
        'Count file size
        FileSize = Val(HextoDec(Hex(f.nFileSizeHigh) + Hex(f.nFileSizeLow)))
    ElseIf Left(f.cFileName, 1) <> "." Then
       FileSize = 0
    End If
    'Close the file search
    FindClose (hFile)
End Function


Function DirSpace(sPath As String) As Double
    Dim f As WIN32_FIND_DATA
    Dim hFile As Long
    Dim hSize As Long
    DirSpace = 0
    'Add the slash to the search path
    sPath = FixPath(sPath)
    hFile = FindFirstFile(sPath & "*.*", f)
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    If (f.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
        'Count file size
        DirSpace = DirSpace + Val(HextoDec(Hex(f.nFileSizeHigh) + Hex(f.nFileSizeLow)))
    ElseIf Left(f.cFileName, 1) <> "." Then
        'call the DirSpace with subdirectory
        DirSpace = DirSpace + DirSpace(FixPath(sPath) & APItoString(f.cFileName))
    End If
    'Enumerate all the files
    Do While FindNextFile(hFile, f)
        If (f.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
            'Count file size
            DirSpace = DirSpace + Val(HextoDec(Hex(f.nFileSizeHigh) + Hex(f.nFileSizeLow)))
        ElseIf Left(f.cFileName, 1) <> "." Then
            'call the DirSpace with subdirectory
            DirSpace = DirSpace + DirSpace(FixPath(sPath) & APItoString(f.cFileName))
        End If
    Loop
    'Close the file search
    FindClose (hFile)
End Function






Function HextoDec(ByVal cadena As String) As Double
Dim i, j, hexnum As Long
Dim salida As Double
Dim caracter, tabla As String
tabla = "123456789ABCDEF"
salida = 0
j = 0
For i = Len(cadena) To 1 Step -1
   j = j + 1
   caracter = Mid(UCase(cadena), i, 1)
   hexnum = InStr(1, tabla, caracter)
   salida = salida + hexnum * 16 ^ (j - 1)
Next i
HextoDec = salida
End Function

Function DecToHex(ByVal numero As Long) As String
Dim cadena As String
cadena = Hex(numero)
Do While Len(cadena) < 5
cadena = "0" + cadena
Loop
DecToHex = cadena
End Function



'Function FileSize(ByVal archivo As String)
'Dim Valor As Double
'Valor = 0
'On Error Resume Next
'Valor = FileLen(archivo)
'On Error GoTo 0
'FileSize = Valor
'End Function


  
Public Sub BorraArchivo(ByVal archivo As String)
 On Error Resume Next
 Kill (archivo)
 On Error GoTo 0
End Sub

' for file size use filelen() function


'Regresa verdadero si el archivo pasado existe, y falso si no.
Function FileExists(ByVal archivo As String) As Boolean
On Error GoTo error1
If Len(archivo) = 0 Then
   FileExists = False
Else
If Dir$(archivo) = "" Then
   FileExists = False
Else
   FileExists = True
End If
End If
On Error GoTo 0
Exit Function
error1:
FileExists = False
On Error GoTo 0
Exit Function
End Function

'Regresa verdadero si el DIRECTORIO pasado existe, y falso si no.
Function DirExists(ByVal DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory

GoTo fin
ErrorHandler:
    DirExists = False
fin:
On Error GoTo 0
End Function

Public Sub copiarchivo(ByVal origen As String, ByVal destino As String)
Const ChunkSize = 16384 / 2
Dim CanalOrigen, CanalDestino As Long
Dim Data As String
  On Error GoTo errores
    CanalOrigen = FreeFile
    Open origen For Binary As #CanalOrigen  'Source
    CanalDestino = FreeFile
    Open destino For Binary As #CanalDestino 'Destination
    If LOF(CanalOrigen) > 0 Then
       Do Until LOF(CanalOrigen) = Loc(CanalOrigen) Or EOF(CanalOrigen)
       'Will do this loop until the end of the source file.
               Data = ""
               If LOF(CanalOrigen) - Loc(CanalOrigen) < ChunkSize Then
                   Data = String(LOF(CanalOrigen) - Loc(CanalOrigen), 0)
               Else
                   Data = String(ChunkSize, 0)
               End If
               Get #CanalOrigen, , Data
               Put #CanalDestino, , Data
       Loop
    Else
       MsgBox ("Archivo " + Chr(10) + origen + Chr(10) + "con longitud igual a cero. Lo voy a saltar")
    End If
    Close #CanalOrigen, #CanalDestino
    On Error GoTo 0
    Exit Sub
errores:
    MsgBox ("Error " + Str(Err.Number) + " " + Err.Description + Chr(13) & Chr(10) + "Accesando el file en " + origen + Chr(13) & Chr(10) + "No puedo leer una de las unidades, espera a que la unidad termine de reconocer el disco y presiona aceptar")
    Resume
    
End Sub

Function RowCount(ByVal archivo As String) As Long
Dim RCanal, Lineas As Long
Dim RLinea As String
RowCount = 0
RCanal = FreeFile
If FileExists(archivo) Then
   Open archivo For Input As #RCanal
   Do While Not EOF(RCanal)
      Line Input #RCanal, RLinea
      If Len(Trim(RLinea)) > 0 Then Lineas = Lineas + 1
   Loop
   Close (RCanal)
End If
RowCount = Lineas
End Function
