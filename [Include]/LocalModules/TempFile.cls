VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTempFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private colTempFiles As New Collection
Public Sub add(sFilename As String)
'Debug.Print colTempFiles.Item(Index)
On Error Resume Next
colTempFiles.Remove (sFilename)
colTempFiles.add sFilename, sFilename
End Sub

Public Sub Remove(Index As Variant)
        colTempFiles.Remove Index
End Sub

Public Function Item(Index As Variant) As String
    Item = colTempFiles.Item(Index)
End Function

Public Function count() As Long
    count = colTempFiles.count
End Function
Public Sub clear()

    Dim tempfile
    'On Error Resume Next
    For Each tempfile In colTempFiles
        Kill tempfile
        RmDir tempfile
        colTempFiles.Remove tempfile
    Next
    
End Sub
Private Sub Class_Terminate()

    Dim tempfile
    On Error Resume Next
    For Each tempfile In colTempFiles
    Kill tempfile
    RmDir tempfile
    Next
    Set colTempFiles = Nothing
End Sub
