VERSION 5.00
Begin VB.Form FrmSelDir 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmSelDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FrmSelDir.Tag = Dir1.Path
FrmSelDir.Visible = False
End Sub

Private Sub Command2_Click()
Dim cadena, dircompleto As String
cadena = InputBox("Nuevo directorio:", "Crear directorio en " + Dir1.Path, "NUEVO")
If cadena <> "" Then
   dircompleto = Dir1.Path + "\" + cadena + "\"
   Call Create_Directory(dircompleto)
   Dir1.Refresh
End If
End Sub

Private Sub Dir1_Change()
If Mid$(Drive1.drive, 1, 2) <> Mid$(Dir1.Path, 1, 2) Then
   Drive1.drive = Dir1.Path
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.drive
End Sub

Sub Create_Directory(sdirectory As String)

'---------------CREATE A DIRECTORY-----------------------
'This procedure creates the directory where the file(s)
'are to be installed. There is some error handling in
'it incase the directory already exists.
'--------------------------------------------------

Dim strpath As String       'The directory which will be created...
Dim intOffset As Integer    'Searches for a "\" so it can create the dirs...
Dim intAnchor As Integer    'Equal to the above variable...
Dim strOldPath As String    'Returns the CurDir to the old path(the dir
                            'the setup file is in)...

On Error Resume Next        'Error handling...

strOldPath = CurDir$        'Find the current Directory...
intAnchor = 0               'Reset intAnchor...

'Searches for the "\" to create the dirs properly...
intOffset = InStr(intAnchor + 1, sdirectory, "\")
intAnchor = intOffset   'Equal to the above...
Do
    intOffset = InStr(intAnchor + 1, sdirectory, "\")
    intAnchor = intOffset
    
    If intAnchor > 0 Then   'If there is 1 or more "\" then...
        
        'Create the directory using the text before the "\"...
        strpath = Left$(sdirectory, intOffset - 1)
        
        ' Determine if this directory already exists...
        Err = 0
        ChDir strpath   'If it does, change to that directory...
        
        If Err Then     'If it doesn't exist...
            
            ' We must create this directory...
            Err = 0
            MkDir strpath   'Make the Directory...
        End If
    End If
Loop Until intAnchor = 0    'Loop until all directories have been made
                            'I.e C:\Prog\David\Cowan is 3 directories...
Done:
    ChDir strOldPath        'Change back to the the 'old' current directory...
Err = 0                     'Reset the error number...
End Sub


