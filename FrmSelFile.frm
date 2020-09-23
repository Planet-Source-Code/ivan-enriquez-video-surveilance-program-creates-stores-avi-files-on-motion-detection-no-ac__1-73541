VERSION 5.00
Begin VB.Form FrmSelFile 
   Caption         =   "Open one or multiple files"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form2"
   ScaleHeight     =   6120
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      Height          =   5745
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   6135
   End
   Begin VB.DirListBox Dir1 
      Height          =   5265
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmSelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
If Mid$(Drive1.drive, 1, 2) <> Mid$(Dir1.Path, 1, 2) Then
   Drive1.drive = Dir1.Path
End If
End Sub

Private Sub Drive1_Change()
On Error GoTo problemas
   Dir1.Path = Drive1.drive
   GoTo salida
problemas:
 MsgBox "Error: cant change to drive " + Drive1.drive
   Resume Next
salida:
On Error GoTo 0
End Sub


Private Sub File1_dblClick()
   FrmSelFile.Tag = File1.Path + "\" + File1.filename
   If InStr(2, FrmSelFile.Tag, "\\") Then FrmSelFile.Tag = Replace(FrmSelFile.Tag, "\\", "\")
   FrmSelFile.Visible = False
End Sub

