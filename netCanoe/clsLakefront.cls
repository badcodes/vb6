VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLakefront"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_taskName As String
Private m_preUrl As String
Private m_sufUrl As String
Private m_dynUrl As String
Private m_savedIn As String

Public Property Get savedIn() As String
    savedIn = m_savedIn
End Property

Public Property Let savedIn(ByVal Value As String)
    m_savedIn = Value
End Property

Public Property Get dynUrl() As String
    dynUrl = m_dynUrl
End Property

Public Property Let dynUrl(ByVal Value As String)
    m_dynUrl = Value
End Property

Public Property Get sufUrl() As String
    sufUrl = m_sufUrl
End Property

Public Property Let sufUrl(ByVal Value As String)
    m_sufUrl = Value
End Property

Public Property Get preUrl() As String
    preUrl = m_preUrl
End Property

Public Property Let preUrl(ByVal Value As String)
    m_preUrl = Value
End Property

Public Property Get taskName() As String
    taskName = m_taskName
End Property

Public Property Let taskName(ByVal Value As String)
    m_taskName = Value
End Property

Public Sub addTask(ByVal sTaskName As String, _
                          ByVal sSavedIN As String, _
                          ByVal sPreUrl As String, _
                          ByVal sDynUrl As String, _
                          ByVal sSufUrl As String _
                          )

End Sub

Public Sub loadCanoe()

End Sub
