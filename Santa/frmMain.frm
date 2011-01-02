VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Begin VB.Form fraMain 
   Caption         =   "Santa cxina"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875.26
   ScaleMode       =   0  'User
   ScaleWidth      =   7012.545
   StartUpPosition =   3  'Windows Default
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
      Height          =   2895
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
      _cx             =   5106
      _cy             =   5106
      FlashVars       =   ""
      Movie           =   "x:\santa.swf"
      Src             =   "x:\santa.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "fraMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim pMovie As String
    pMovie = Environ$("temp") & "\cxina_santa.swf"
    Dim pImg As String
    pImg = Environ$("temp") & "\cxina_santa.jpg"
    ExtractMovie pMovie, 101
    ExtractJpg pImg, 101
    flash.Movie = pMovie
    flash.FlashVars = "user=4111899&photo=cxina_santa.jpg&photo_scale=72&photo_rotate=-16&photo_x=8&photo_y=5&message=" & LoadResString(101)
    flash.Play
End Sub

Private Sub ExtractResource(ByVal vDst As String, ByVal vType As String, ByVal vId As Long)
        '<EhHeader>
        On Error GoTo ExtractMovie_Err
        '</EhHeader>
        Dim pData() As Byte
100     If FileExists(vDst) Then Exit Sub
102     pData = LoadResData(vId, vType)
        Dim fNum As Integer
104     fNum = FreeFile
106     Open vDst For Binary Access Write As #fNum
108     Put #fNum, , pData
110     Close #fNum
        '<EhFooter>
        Exit Sub

ExtractMovie_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Project1.Form1.ExtractMovie " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub
'CSEH: ErrMsgBox
Private Sub ExtractMovie(ByVal vDst As String, ByVal vId As Long)
    ExtractResource vDst, "FLASH", vId
End Sub

Private Sub ExtractJpg(ByVal vDst As String, ByVal vId As Long)
    ExtractResource vDst, "JPEG", vId
End Sub


Private Function FileExists(ByRef vPath As String) As Boolean
    On Error Resume Next
    Err.Clear
    FileExists = True
    Call FileLen(vPath)
    If Err.Number <> 0 Then FileExists = False
End Function

Private Sub Form_Resize()
    On Error Resume Next
    flash.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
