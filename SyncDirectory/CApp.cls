VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IApp
Const c_APPNAME As String = "SyncDirectory"

Private m_frmMain As Form

Private Sub IApp_Initialize(MainForm As Form)
    Set m_frmMain = MainForm
End Sub

Private Property Set IApp_MainForm(RHS As Form)
    Set m_frmMain = RHS
End Property

Private Property Get IApp_MainForm() As Form
    Set iapp_mainfrom = m_frmMain
End Property

Private Sub IApp_OnLoad()
End Sub

Private Sub IApp_OnUnload()
    Set m_frmMain = Nothing
End Sub

Private Property Get IApp_Title() As String
    IApp_Title = c_APPNAME
End Property


