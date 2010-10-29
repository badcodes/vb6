VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFolderBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHwnd As Long
Private mDir As String
Private mTitle As String
Private Const MAX_PATH As Long = 260
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)
    
Public Property Get Owner() As Long
    Owner = mHwnd
End Property

Public Property Let Owner(ByVal hwnd As Long)
    mHwnd = hwnd
End Property

Public Property Get InitDirectory() As String
    InitDirectory = mDir
End Property

Public Property Let InitDirectory(ByVal sPath As String)
    mDir = sPath
End Property

Public Property Get Title() As String
    Title = mTitle
End Property

Public Property Let Title(ByVal sTitle As String)
    mTitle = sTitle
End Property
Public Function Browse() As String

   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim lpSelPath As Long
   Dim sPath As String * MAX_PATH
   
   With BI
      .hOwner = mHwnd
      .pidlRoot = 0
      .lpszTitle = mTitle
      .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
      lpSelPath = LocalAlloc(LPTR, LenB(mDir) + 1)
      CopyMemory ByVal lpSelPath, ByVal mDir, LenB(mDir)
      .lParam = lpSelPath
   End With
    
   pidl = SHBrowseForFolder(BI)
   
   If pidl Then
     
      If SHGetPathFromIDList(pidl, sPath) Then
         Browse = Left$(sPath, InStr(sPath, vbNullChar) - 1)
      Else
         Browse = vbNullString
      End If
      
      Call CoTaskMemFree(pidl)
   
   Else
      Browse = vbNullString
   End If
   
  Call LocalFree(lpSelPath)

End Function

Private Sub Class_Initialize()
    mTitle = "Select A Folder:"
    mDir = vbNullString
End Sub
