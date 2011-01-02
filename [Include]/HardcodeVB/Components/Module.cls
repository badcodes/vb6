VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorModule
    eeBaseModule = 13130    ' CModule
End Enum

Private idProcess As Long
Private hModule As Long
Private sExePath As String

Sub Create(idProcessA As Long, hModuleA As Long, sExePathA As String)
    idProcess = idProcessA
    hModule = hModuleA
    sExePath = sExePathA
End Sub

Property Get Handle() As Long
    Handle = hModule
End Property

Property Get ProcessID() As Long
    ProcessID = idProcess
End Property

Property Get ExePath() As String
    ExePath = sExePath
End Property

Property Get ExeFile() As String
    ExeFile = MUtility.GetFileBaseExt(sExePath)
End Property

Property Get ExeName() As String
    ExeName = MUtility.GetFileBase(sExePath)
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Module"
        Select Case e
        Case eeBaseModule
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

