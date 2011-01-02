VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWinHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let Url(ByVal vUrl As String)

End Property

Public Sub ReportError(ByVal vTag As Variant, ByVal vErrorNum As Long, ByVal vErrorDesc As String)

End Sub

Public Sub ReportDataAvail(ByVal vTag As Variant, ByRef vData() As Byte, ByVal vTotalByte As Long)

End Sub

Public Sub ReportRespone(ByVal vTag As Variant, ByVal vStatus As Long, ByVal vStatusText As String)

End Sub

Public Sub ReportComplete(ByVal vTag As Variant, vStatus As Long)

End Sub

Public Sub ReportProgress(ByVal vTag As Variant, ByVal vMin As Long, ByVal vMax As Long)

End Sub

