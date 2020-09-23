Attribute VB_Name = "Module1"
Global NoiseReductionLevel As Byte
Global MotionTriggerLevel As Long
Global ScanGridDensity As Long
Global FileLenSegundos As Long
Global AviFps As Long
Global Turno As Long
Global CurrentFilename As String
Global ScanFreq As Long
Global Ancho As Long, Alto As Long
Global MinutosSplit As Long
Global FIleDIr As String
Global ArchivoEnTurno As String
Global TimeIni As Date
Global CompresionOptionsDone As Boolean
Global Dispositivo As Integer
Global SecondsLeft As Long
Global SecondsLen As Long
Global BeepOnDetection As Long
Sub Main()

If Val(Command$) > 0 Then
   Dispositivo = Val(Command$)
Else
   Dispositivo = 0
End If
FileLenSegundos = 5

Form1.Show

End Sub


