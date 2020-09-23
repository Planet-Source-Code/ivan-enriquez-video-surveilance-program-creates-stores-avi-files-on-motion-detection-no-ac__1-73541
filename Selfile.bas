Attribute VB_Name = "Selfile"
Public Function selectfile(ByVal DefaultDir As String, ByVal DefaultExt As String) As String
FrmSelDir.Tag = ""
FrmSelFile.Tag = ""
If DefaultDir <> "" Then
   FrmSelFile.Dir1.Path = DefaultDir
   FrmSelDir.Tag = DefaultDir
End If
If DefaultExt <> "" Then
   FrmSelFile.File1.Pattern = DefaultExt
   FrmSelFile.Tag = DefaultExt
End If
FrmSelFile.Show 1
selectfile = FrmSelFile.Tag
Unload FrmSelFile
End Function

Public Function SelectDir(ByVal DefaultDir As String) As String
If DefaultDir <> "" Then
   If DirExists(DefaultDir) Then
      FrmSelDir.Drive1.Drive = Mid(DefaultDir, 1, 2)
      FrmSelDir.Dir1.Path = DefaultDir
   Else
      FrmSelDir.Dir1.Path = App.Path
   End If
End If
FrmSelDir.Show 1
SelectDir = FrmSelDir.Tag
Unload FrmSelDir
End Function

