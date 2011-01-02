VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDriveWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Implement Basic-friendly version of IEnumVARIANT
Implements IVariantWalker
' Connect back to parent collection
Private connect As CDrives

Public Enum EErrorDriveWalker
    eeBaseDriveWalker = 13040   ' CDriveWalker
End Enum

' Private state data
Private i As Long

Private Sub Class_Initialize()
    ' Initialize position in collection
    i = 1
    BugLocalMessage "CDriveWalker:Initialize"
End Sub

Private Sub Class_Terminate()
    BugLocalMessage "CDriveWalker:Terminate"
End Sub

' Receive connection from CDrives
Friend Sub Attach(connectA As CDrives)
    Set connect = connectA
End Sub

' Return IEnumVARIANT (indirectly) to client collection
Friend Property Get NewEnum() As IEnumVARIANT
    ' Delegate to class that implements real IEnumVARIANT
    Dim vars As CEnumVariant
    ' Connect walker to CEnumVariant so it can call methods
    Set vars = New CEnumVariant
    vars.Attach Me
    ' Return walker to collection data
    Set NewEnum = vars
End Property

' Implement IVariantWalker methods
Private Function IVariantWalker_More(v As Variant) As Boolean
    ' Find the next drive and return it through reference
    Do While i <= 26
        ' Check flags to see if next drive exists
        If MBytes.RShiftDWord(connect.DriveFlags, i - 1) And 1 Then
            Dim drive As CDrive
            Set drive = New CDrive
            drive.Root = i
            Set v = drive
            IVariantWalker_More = True
            i = i + 1
            Exit Function
        End If
        i = i + 1
    Loop
End Function

Private Sub IVariantWalker_Skip(c As Long)
    ' Skip ahead in the iteration
    i = i + c
End Sub

Private Sub IVariantWalker_Reset()
    ' Reset the iteration
    i = 1
End Sub
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".DriveWalker"
        Select Case e
        Case eeBaseDriveWalker
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


