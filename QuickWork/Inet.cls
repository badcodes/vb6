VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents myInet As InetCtlsObjects.Inet
Attribute myInet.VB_VarHelpID = -1

Public Sub Init(ByRef vInet As Inet)
    Set myInet = vInet
    myInet.Execute "http://www.planet-source-code.com/upload_PSC/ftp/PSC_Browse15136712152002.zip", "GET"
End Sub

Private Sub myInet_StateChanged(ByVal State As Integer)
    Debug.Print State
End Sub
