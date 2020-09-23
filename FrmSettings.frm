VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSettings 
   Caption         =   "Settings"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   LinkTopic       =   "Form2"
   ScaleHeight     =   5070
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Save and close"
      Height          =   375
      Left            =   5640
      TabIndex        =   18
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   3480
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Beep on motion detection"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   3000
      Width           =   255
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1085
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   1
      Max             =   150
      SelStart        =   15
      TickStyle       =   1
      TickFrequency   =   5
      Value           =   15
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      _Version        =   393216
      Min             =   1
      Max             =   1000
      SelStart        =   1
      TickStyle       =   1
      TickFrequency   =   50
      Value           =   1
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      _Version        =   393216
      Min             =   1
      Max             =   5
      SelStart        =   2
      TickStyle       =   1
      Value           =   2
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   1920
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   2000
      SelStart        =   50
      TickFrequency   =   50
      Value           =   50
   End
   Begin VB.Label Label10 
      Caption         =   "Minutes to split video files:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Avi FPS:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Time to record after last motion(seconds)"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label8 
      Caption         =   "Path for video files:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Poll freq (ms)"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Noise Reduction:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "MotionTrigger Level:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Scan grid Density:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
   BeepOnDetection = Check1.Value
End Sub

Private Sub Command1_Click()
Dim cadena As String
cadena = SelectDir(FIleDIr)
If cadena <> "" Then
   FIleDIr = cadena
   Label9.Caption = FIleDIr
End If
End Sub

Private Sub Command2_Click()
SaveSetting "MotionDetect", "Param", "NoiseReductionLevel", Str(NoiseReductionLevel)
SaveSetting "MotionDetect", "Param", "MotionTriggerLevel", Str(MotionTriggerLevel)
SaveSetting "MotionDetect", "Param", "ScanGridDensity", Str(ScanGridDensity)
SaveSetting "MotionDetect", "Param", "SecondsLen", Str(SecondsLen)
SaveSetting "MotionDetect", "Param", "Turno", Str(Turno)
SaveSetting "MotionDetect", "Param", "FileLenSegundos", Str(FileLenSegundos)
SaveSetting "MotionDetect", "Param", "ScanFreq", Str(ScanFreq)
SaveSetting "MotionDetect", "Param", "FileDir", FIleDIr
SaveSetting "MotionDetect", "Param", "MinutosSplit", Str(MinutosSplit)
SaveSetting "MotionDetect", "Param", "AviFps", Str(AviFps)
SaveSetting "MotionDetect", "Param", "BeepOnDetection", Str(BeepOnDetection)
FrmSettings.Visible = False
End Sub

Private Sub Form_Load()
Slider1.Value = NoiseReductionLevel
Slider2.Value = MotionTriggerLevel
Slider3.Value = ScanGridDensity
Slider4.Value = ScanFreq
Text1.Text = SecondsLen
Text2.Text = MinutosSplit
Text3.Text = AviFps
Label9.Caption = FIleDIr
Check1.Value = BeepOnDetection
End Sub

Private Sub Slider1_Click()
NoiseReductionLevel = Slider1.Value
End Sub

Private Sub Slider2_Click()
MotionTriggerLevel = Slider2.Value
Call Progress(Form1.Picture6, MotionTriggerLevel, 1000, vbGreen, vbRed)
End Sub

Private Sub Slider3_Click()
ScanGridDensity = Slider3.Value
Form1.Picture1.Picture = Nothing
End Sub

Private Sub Slider4_Click()
ScanFreq = Val(Slider4.Value)
Form1.Timer1.Interval = ScanFreq
End Sub

Private Sub Text1_Change()
SecondsLen = Val(Text1.Text)
End Sub

Private Sub Text2_Change()
If Val(Text2.Text) > 0 Then
   MinutosSplit = Val(Text2.Text)
End If
End Sub

Private Sub Text3_Change()
If Val(Text3.Text) > 0 Then
   AviFps = Val(Text3.Text)
End If
End Sub
