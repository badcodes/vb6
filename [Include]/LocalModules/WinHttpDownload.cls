VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinHttpDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents WinHttpConnection As CWinHTTP
Attribute WinHttpConnection.VB_VarHelpID = -1
Private m_ID As Long

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let ID(ByVal LonValue As Long)
    m_ID = LonValue
End Property

Private Sub Class_Initialize()
    Set WinHttpConnection = New CWinHTTP
End Sub

