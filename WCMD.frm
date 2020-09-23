VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WCMD [WebCam Motion Detector]"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13380
   Icon            =   "WCMD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   13380
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   11640
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   8
      Top             =   1080
      Width           =   4815
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   840
      ScaleHeight     =   105
      ScaleWidth      =   8385
      TabIndex        =   3
      Top             =   480
      Width           =   8415
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   9240
      Top             =   4560
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   9105
      TabIndex        =   2
      Top             =   0
      Width           =   9135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3720
      Left            =   120
      ScaleHeight     =   248
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   328
      TabIndex        =   0
      Top             =   720
      Width           =   4920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8520
      Top             =   4560
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   105
      ScaleWidth      =   9105
      TabIndex        =   5
      Top             =   240
      Width           =   9135
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   1920
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   9360
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   9360
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Menu Startcam 
      Caption         =   "StartCam"
   End
   Begin VB.Menu StopCam 
      Caption         =   "StopCam"
   End
   Begin VB.Menu Sensibility 
      Caption         =   "Settings"
   End
   Begin VB.Menu VSS 
      Caption         =   "Video Settings"
      Begin VB.Menu VF 
         Caption         =   "Video Format"
      End
      Begin VB.Menu VD 
         Caption         =   "Video Display"
      End
      Begin VB.Menu VS 
         Caption         =   "Video Source"
      End
   End
   Begin VB.Menu st 
      Caption         =   "Start"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnstop 
      Caption         =   "Stop"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function TextOut& Lib "gdi32" Alias "TextOutA" _
(ByVal HDC&, ByVal X&, ByVal y&, ByVal lpString$, ByVal nCount&)
Private Declare Function SetBkMode& Lib "gdi32" (ByVal HDC&, _
ByVal nBkMode&)
Private Declare Function SetTextColor& Lib "gdi32" (ByVal HDC&, ByVal Cl&)
Private Declare Function Beep Lib "kernel32" _
  (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public mCapHwnd As Long
Dim sFileName As String
Dim Capturando As Boolean
Dim bmp As cDIBSection
Dim FramesOnThisAvi As Double
    Dim msgString As String
    Dim bmpFile As String
    Dim sFile As String
    Dim res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim BI As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim I As Long
    Dim CurrentNewFrame As Long

Dim Caps As CAPDRIVERCAPS
Dim CAP_PARAMS As CAPTUREPARMS

Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long

Private Sub Form_Load()
loguea "Program start"
'recover saved values (or set them for first time)
CompresionOptionsDone = False
NoiseReductionLevel = GetSetting("MotionDetect", "Param", "NoiseReductionLevel", "50")
MotionTriggerLevel = GetSetting("MotionDetect", "Param", "MotionTriggerLevel", "50")
ScanGridDensity = GetSetting("MotionDetect", "Param", "ScanGridDensity", "1")
SecondsLen = Val(GetSetting("MotionDetect", "Param", "SecondsLen", "5"))
Turno = Val(GetSetting("MotionDetect", "Param", "Turno", "1"))
FileLenSegundos = Val(GetSetting("MotionDetect", "Param", "FileLenSegundos", "600"))
ScanFreq = Val(GetSetting("MotionDetect", "Param", "ScanFreq", "250"))
FIleDIr = GetSetting("MotionDetect", "Param", "FileDir", FixPath(App.Path))
MinutosSplit = Val(GetSetting("MotionDetect", "Param", "MinutosSplit", "60"))
AviFps = Val(GetSetting("MotionDetect", "Param", "AviFps", "10"))
BeepOnDetection = GetSetting("MotionDetect", "Param", "BeepOnDetection", "1")

Timer1.Interval = ScanFreq





Set bmp = New cDIBSection
CurrentFilename = GetCurrentFile()
End Sub

Private Sub Adjust_Click()
FrmSettings.Visible = True
End Sub



Private Sub TomaTamaño()
Picture1.Width = Ancho * Screen.TwipsPerPixelX
Picture1.Height = Alto * Screen.TwipsPerPixelY
Picture1.left = 20
Picture2.top = Picture1.top
Picture2.left = Picture1.left + Picture1.Width
Picture2.Width = Ancho * Screen.TwipsPerPixelX
Picture2.Height = Alto * Screen.TwipsPerPixelY
Picture1.Picture = Nothing
Picture2.Picture = Nothing


Form1.Width = Picture2.left + Picture2.Width + 40
Form1.Height = Picture1.top + Picture1.Height + 1000
Label4.left = Form1.Width - Label4.Width - 200
Label4.top = Picture3.top
Label6.left = Form1.Width - Label6.Width - 200
Label6.top = Picture4.top
Call Progress(Picture6, MotionTriggerLevel, 1000, vbGreen, vbRed)
End Sub



Function GetCurrentFile() As String
Turno = Turno + 1
If Turno = 10000 Then Turno = 1
GetCurrentFile = Format(Turno, "0000") + "_" + Format(Now, "yyyymmdd_hhmmss") + ".avi"
SaveSetting "MotionDetect", "Param", "Turno", Str(Turno)
End Function

Private Sub Form_Unload(Cancel As Integer)
   StopCamx
   End
End Sub


Private Sub Command3_Click()
'If capFileSetCaptureFile(mCapHwnd, CurrentFilename) Then MsgBox ("error")
SendMessage mCapHwnd, WM_CAP_SINGLE_FRAME_OPEN, 0, 0
End Sub

'Private Sub Command4_Click()
'temp = SendMessage(mCapHwnd, WM_CAP_DLG_VIDEOCONFIG, 0&, 0&)
'        DoEvents
'End Sub



Private Sub mnstop_Click()
Capturando = False

    If (ps <> 0) Then Call AVIStreamClose(ps)
    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)
    If (pfile <> 0) Then Call AVIFileClose(pfile)
    Call AVIFileExit
    If (res <> AVIERR_OK) Then
        If res = AVIERR_BADFORMAT Or res = AVIERR_INTERNAL Then
            MsgBox "There was an error creating the AVI File." + vbCrLf + "Probably the choosen Video Compression does not support the File format or the input File is corrupt", vbInformation, App.Title
        Else
            MsgBox "There was an error creating the AVI File.", vbInformation, App.Title
        End If
    Else
       DoEvents
       loguea "Se cerró el archivo " + sFileName + " con " + Str(FramesOnThisAvi) + " frames"
       If FramesOnThisAvi < 10 Then
          Call BorraArchivo(sFileName)
          loguea "Se borro el archivo debido a que tenia menos de 10 frames"
       End If
    End If
    
mnstop.Enabled = False
StopCam.Enabled = True
st.Enabled = True
End Sub


Private Sub Sensibility_Click()
FrmSettings.Show
End Sub

Private Sub st_Click()
st.Enabled = False
StopCam.Enabled = False
Call StartCap
mnstop.Enabled = True
End Sub

Private Sub Startcam_Click()
Startcamx
Picture1.AutoRedraw = True
Picture2.AutoRedraw = True

Timer1.Enabled = True
mnstop.Enabled = False
st.Enabled = True
VSS.Enabled = True
End Sub

Private Sub StopCam_Click()
StopCamx
Picture1.Picture = Nothing
Picture2.Picture = Nothing
VSS.Enabled = False
st.Enabled = False
mnstop.Enabled = False
End Sub


Public Function DattedLinex(ByVal cadena As String) As String
DattedLinex = Format(Now, "yyyymmmdd hh:mm:ss") + " " + cadena
End Function

Private Sub Timer1_Timer()
Dim Movim As Long
'Form1.Caption = Str(Timer1.Interval)
'getting picture from camera
Form1.Caption = DattedLinex(sFileName)
'SendMessage mCapHwnd, GET_FRAME, 0, 0
capGrabFrameNoStop (mCapHwnd)
SendMessage mCapHwnd, COPY, 0, 0
On Error Resume Next
Picture1.Picture = Clipboard.GetData: Clipboard.Clear
On Error GoTo 0
Movim = 0

Movim = ImageDifference(Picture1, Picture2)

Label4.Caption = Movim
If Movim < MotionTriggerLevel Then
   Call Progress(Picture3, Minimo(Movim, 1000), 1000, vbGreen)
Else
   Call Progress(Picture3, Minimo(Movim, 1000), 1000, vbRed)
   If BeepOnDetection Then Beep Minimo(32760, Maximo(32, Movim)), 10
   SecondsLeft = SecondsLen
End If



Picture2.Picture = Picture1.Picture

If SecondsLeft > 0 Then
    If Capturando Then
       bmp.CreateFromPicture Picture2
       DoEvents
       SetBkMode bmp.HDC, 1
       SetTextColor bmp.HDC, vbBlack
       TextOut bmp.HDC, 0, 0, Format(Now, "yyyy mmm dd hh:mm:ss"), Len(Format(Now, "yyyy mmm dd hh:mm:ss"))
       SetTextColor bmp.HDC, vbWhite
       TextOut bmp.HDC, 2, 2, Format(Now, "yyyy mmm dd hh:mm:ss"), Len(Format(Now, "yyyy mmm dd hh:mm:ss"))
       CurrentNewFrame = CurrentNewFrame + 1
       res = AVIStreamWrite(psCompressed, CurrentNewFrame, 1, bmp.DIBSectionBitsPtr, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
       FramesOnThisAvi = FramesOnThisAvi + 1
       Label12.Caption = Str(FramesOnThisAvi)
       If res <> AVIERR_OK Then MsgBox ("error")
   End If
End If


If DateDiff("n", TimeIni, Now) >= MinutosSplit Then  'hay que cambiar de AVI?
   If Capturando Then
      Call mnstop_Click 'para la grabacion
      DoEvents
      Call st_Click 'comienza la grabacion en un nuevo AVI
   End If
End If

End Sub

Sub StopCamx()
Call AVIFileExit
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
Timer1.Enabled = False

End Sub

Public Sub GetCapss()
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    
    'Get the capture window attributes
    retVal = capGetStatus(mCapHwnd, capStat)
        
    If retVal Then
        'Resize the main form to fit
       Ancho = capStat.uiImageWidth
       Alto = capStat.uiImageHeight
    Else
       MsgBox ("ooopss no cap status")
    End If
End Sub

Sub Startcamx()

Dim lpszName As String * 100
Dim lpszVer As String * 100
Dim Caps As CAPDRIVERCAPS
        
'//Create Capture Window
capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
mCapHwnd = capCreateCaptureWindowA(lpszName, 0, 0, 0, Ancho, Alto, Me.hWnd, 0)
loguea ("pase getdriverdescription")
'// Set title of window to name of driver
'Form1.Caption = lpszName
capSetCallbackOnStatus mCapHwnd, AddressOf MyStatusCallback
capSetCallbackOnError mCapHwnd, AddressOf MyErrorCallback
'Getting handle of camera window
'mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, Ancho, Alto, Me.hWnd, 0)

If capDriverConnect(mCapHwnd, 0) Then
   loguea ("pase capdriverconnect")
   '/////
   '// Only do the following if the connect was successful.
   '// if it fails, the error will be reported in the call
   '// back function.
   '/////
   '// Get the capabilities of the capture driver
   capDriverGetCaps mCapHwnd, VarPtr(Caps), Len(Caps)
   loguea ("pase capdrivercaps")
   '// If the capture driver does not support a dialog, grey it out
   '// in the menu bar.
   VS.Enabled = False
   VF.Enabled = False
   VF.Enabled = False
   If Caps.fHasDlgVideoSource = 1 Then VS.Enabled = True
   If Caps.fHasDlgVideoFormat = 1 Then VF.Enabled = True
   If Caps.fHasDlgVideoDisplay = 1 Then VD.Enabled = True

   '// Turn Scale on
   'capPreviewScale mCapHwnd, False
            
   '// Set the preview rate in milliseconds
   'capPreviewRate mCapHwnd, 100
        
   '// Start previewing the image from the camera
   capPreview mCapHwnd, False
Else
   MsgBox ("Fallo capdriverconnect")
End If


Call GetCapss
Call TomaTamaño

End Sub

Private Sub Timer2_Timer()
Label6.Caption = SecondsLeft
Call Progress(Picture4, SecondsLeft, SecondsLen, vbBlue, vbWhite)
If SecondsLeft > 0 Then
   SecondsLeft = SecondsLeft - 1
End If
DoEvents
End Sub

Private Sub VF_Click()
Dim temp As Long
      temp = SendMessage(mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
            DoEvents
Call GetCapss
Call TomaTamaño

End Sub

Private Sub VS_Click()
 capDlgVideoSource mCapHwnd

End Sub

Private Sub StartCap()
  FramesOnThisAvi = 0
  CurrentNewFrame = 0
  TimeIni = Now
  sFileName = FixPath(FIleDIr) + GetCurrentFile()
  res = AVIFileOpen(pfile, sFileName, OF_WRITE Or OF_CREATE, 0&) 'inicializa avi
    If (res <> AVIERR_OK) Then MsgBox Error
    loguea "New file: " + sFileName

    'obtiene un bmp de la camara para sacar algunos parametros a especificar para el avi
    Set bmp = New cDIBSection
    bmp.CreateFromPicture Picture2
'   Llena datos del encabezado el avi
    
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)            '// stream type video
        .fccHandler = 0&                                     '// default AVI handler
        .dwScale = 1
        .dwRate = AviFps                                '// fps
        .dwSuggestedBufferSize = bmp.Width * bmp.Height      '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)  '// rectangle for stream
    End With
    
    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

'   Crea el stream original para el avi
'   pfile=pointer to the filename
'   ps=pointer to the avi stream
'   strhdr=estructura de encabezado

    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then MsgBox Error

    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    'El pedido de compresion settings solo se hara una vez por programa
    If CompresionOptionsDone = False Then
       pOpts = VarPtr(opts)
       res = AVISaveOptions(Form1.hWnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)
       'returns TRUE if User presses OK, FALSE if Cancel, or error code
       If res <> 1 Then 'In C TRUE = 1
           Call AVISaveOptionsFree(1, pOpts)
           MsgBox Error
       End If
       CompresionOptionsDone = True
    End If
    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then MsgBox Error
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then MsgBox Error

Capturando = True
End Sub
