VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EErrorAbout
    eeBaseAbout = 13660     ' CAbout
    eeAppNotInit            ' Client must be initialized with App object
End Enum

' Private form variable
Private about As New FAbout

' Could be App object, or anything with the same version properties
Public Client As Object
' Normally the icon of the client application
Public Icon As Picture
' Miscellaneous properties
Public InfoProg As String
Public Copyright As String
Public Comments As String
Public SecretButton As Integer
Public SecretKey As Integer
Public SecretShift As Integer
Public Animator As IAnimation

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".CAbout"
        Select Case e
        Case eeBaseAbout
            BugAssert True
        Case eeAppNotInit
            sText = "Load: Client must be initialized with App object"
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If
End Sub
#End If

Sub Load()
With about
    ' We need version properties to display on About form
    If Client Is Nothing Then ErrRaise eeAppNotInit
    ' Pass other optional properties through to form
    Set .Client = Client
    Set .ClientIcon = Icon
    .InfoProg = InfoProg
    .Copyright = Copyright
    .Comments = Comments
    .SecretButton = SecretButton
    .SecretKey = SecretKey
    .SecretShift = SecretShift
    .UserInfo(1) = UserInfo(1)
    .UserInfo(2) = UserInfo(2)
    .UserInfo(3) = UserInfo(3)
    ' Show the form
    .Show vbModal
End With
End Sub

Property Get UserInfo(i As Integer) As String
    UserInfo = about.UserInfo(i)
End Property

Property Let UserInfo(i As Integer, asUserInfoA As String)
    about.UserInfo(i) = asUserInfoA
End Property

Private Sub Class_Initialize()
    Debug.Print "Initializing CAbout"
End Sub

Private Sub Class_Terminate()
    Debug.Print "Terminating CAbout"
End Sub
