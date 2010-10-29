VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'$ Uses DRIVE.CLS

' Private version of data structure
Private af As Long

Public Enum eeErrorDrives
    eeBaseDrives = 13030    ' CDrives
End Enum

Private Sub Class_Initialize()
    ' Initialize internal data
    af = GetLogicalDrives()
    BugLocalMessage "CDrives:Initialize"
End Sub

Private Sub Class_Terminate()
    BugLocalMessage "CDrives:Terminate"
End Sub

' Friend properties to make data accessible to data walker class
Friend Property Get DriveFlags() As Long
    DriveFlags = af
End Property

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
' Create a new data walker object and connect to it
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    ' Create a new iterator object
    Dim drivewalker As CDriveWalker
    Set drivewalker = New CDriveWalker
    ' Connect it with collection data
    drivewalker.Attach Me
    ' Return it
    Set NewEnum = drivewalker.NewEnum
End Function

Public Property Get Count() As Integer
    Dim c As Long, i As Long
    For i = 0 To 25
        If MBytes.RShiftDWord(af, i) And &H1 Then c = c + 1
    Next
    Count = c
End Property

' Default property
Public Property Get item(v As Variant) As CDrive
Attribute item.VB_UserMemId = 0
    ' Return default (Nothing) if error
    Dim drive As CDrive
    Set drive = New CDrive
    drive.Root = v
    Set item = drive
End Property
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Drives"
        Select Case e
        Case eeBaseDrive
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

