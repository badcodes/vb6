VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDownloadProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_StrURL As String
Private m_LonCurrentBytes As Long
Private m_LonTotalBytes As Long
Private m_StrTextInfo As String
Private m_bEnable As Boolean
Private m_StrSaveAS As String

Public Property Get saveAs() As String
    saveAs = m_StrSaveAS
End Property

Public Property Let saveAs(ByVal StrValue As String)
    m_StrSaveAS = StrValue
End Property

Public Property Get Enable() As Boolean
    Enable = m_bEnable
End Property

Public Property Let Enable(ByVal bValue As Boolean)
    m_bEnable = bValue
End Property

Public Property Get TextInfo() As String
    TextInfo = m_StrTextInfo
End Property

Public Property Let TextInfo(ByVal StrValue As String)
    m_StrTextInfo = StrValue
End Property

Public Property Get TotalBytes() As Long
    TotalBytes = m_LonTotalBytes
End Property

Public Property Let TotalBytes(ByVal LonValue As Long)
    m_LonTotalBytes = LonValue
End Property

Public Property Get CurrentBytes() As Long
    CurrentBytes = m_LonCurrentBytes
End Property

Public Property Let CurrentBytes(ByVal LonValue As Long)
    m_LonCurrentBytes = LonValue
End Property

Public Property Get URL() As String
    URL = m_StrURL
End Property

Public Property Let URL(ByVal StrValue As String)
    m_StrURL = StrValue
End Property
