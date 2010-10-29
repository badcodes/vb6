Attribute VB_Name = "MFileExecutor"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const ERROR_BAD_FORMAT = 11&
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const SE_ERR_ACCESSDENIED = 5
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_FNF = 2
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_OOM = 8
Public Const SE_ERR_PNF = 3
Public Const SE_ERR_SHARE = 26

Public Sub FileExecutor(lhWnd As Long, Path As String, Action As String, Optional cParms As Variant, Optional nShowCmd As Variant)

Dim lRtn As Long 'declare the needed variables

lRtn = ShellExecute(lhWnd, Action, Path, 0&, Path, SW_NORMAL) 'execute or print the file or folder
    
If lRtn <= 32 Then 'if an error is found then call the FileError function
   FileError (lRtn)
End If

End Sub


Public Sub FileError(lRtn As Long)
   
Dim Msg As String
    
Select Case lRtn 'if any errors occur then display them to the user
       Case 0
       Msg = "Memory Error"
       Case ERROR_BAD_FORMAT
       Msg = "Bad Executeable Format"
       Case ERROR_FILE_NOT_FOUND
       Msg = "File not found"
       Case ERROR_PATH_NOT_FOUND
       Msg = "Path not found"
       Case SE_ERR_ACCESSDENIED
       Msg = "Access Denied"
       Case SE_ERR_ASSOCINCOMPLETE
       Msg = "Association incomplete"
       Case SE_ERR_DDEBUSY
       Msg = "DDE Busy error"
       Case SE_ERR_DDEFAIL
       Msg = "DDE failed"
       Case SE_ERR_DDETIMEOUT
       Msg = "DEE time out"
       Case SE_ERR_FNF
       Msg = "File not found"
       Case SE_ERR_NOASSOC
       Msg = "No association for this file"
       Case SE_ERR_OOM
       Msg = "Out of Memory"
       Case SE_ERR_PNF
       Msg = "Path could not be found"
       Case SE_ERR_SHARE
       Msg = "Sharing violation"
       Case Else
       Msg = "Unknown Error!, Please try again..."
End Select
    
MsgBox Msg, vbCritical
    
End Sub
