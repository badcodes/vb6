VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDrivesO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'$ Uses DRIVE.CLS

Private drives As New Collection

Private Sub Class_Initialize()
    Refresh
End Sub

' Argument handy for refreshing local and/or remote, but not floppies
Public Sub Refresh(Optional iFirst As Integer = 1)
    Dim i As Integer, af As Long, sRoot As String
    Dim drive As CDrive
    ' Remove old ones
    Do While drives.Count > iFirst
        drives.Remove iFirst
    Loop
    ' Insert new
    af = GetLogicalDrives()
    For i = iFirst To 26
        If RShiftDWord(af, i - 1) And 1 Then
            Set drive = New CDrive
            drive.Root = i
            drives.Add drive, drive.Root
        End If
    Next
End Sub

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = drives.[_NewEnum]
End Function

Public Property Get Count() As Integer
    Count = drives.Count
End Property

' Default property
Public Property Get item(v As Variant) As CDrive
Attribute item.VB_UserMemId = 0
    ' Return default (Nothing) if error
    On Error Resume Next
    Set item = drives(v)
End Property

