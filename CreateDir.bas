Attribute VB_Name = "CreateDir"
Option Explicit

Sub Create_Directory(ByVal sdirectory As String)

'---------------CREATE A DIRECTORY-----------------------
'This procedure creates the directory where the file(s)
'are to be installed. There is some error handling in
'it incase the directory already exists.
'--------------------------------------------------

Dim strPath As String       'The directory which will be created...
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
        strPath = Left$(sdirectory, intOffset - 1)
        
        ' Determine if this directory already exists...
        Err = 0
        ChDir strPath   'If it does, change to that directory...
        
        If Err Then     'If it doesn't exist...
            
            ' We must create this directory...
            Err = 0
            MkDir strPath   'Make the Directory...
        End If
    End If
Loop Until intAnchor = 0    'Loop until all directories have been made
                            'I.e C:\Prog\David\Cowan is 3 directories...
Done:
    ChDir strOldPath        'Change back to the the 'old' current directory...
Err = 0                     'Reset the error number...
End Sub

