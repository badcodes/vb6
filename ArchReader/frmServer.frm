VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "zhServer"
   ClientHeight    =   3096
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   4680
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3096
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock sck 
      Index           =   0
      Left            =   660
      Top             =   450
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim hssZh As HttpServerSet

Dim sHSHead As String


Private Sub Form_Load()


'Dim hSetting As New CSetting
'hSetting.iniFile = MainFrm.zhtmIni
'With hSetting
'.Load Me, SF_Tag
'.Load Me, SF_CAPTION
'End With
'Set hSetting = Nothing

'MZhReaderViaHttpServer.getServerSetting iniHssZh, hssZh
MainFrm.Icon = Me.Icon

'sck(0).LocalPort = CLngStr(Me.Tag)  ' hssZh.sPort)

'If App.PrevInstance = False Or sck(0).LocalPort = 0 Then
    sck(0).Listen
    DoEvents
    'Me.Tag = CStr(sck(0).LocalPort)
'End If

'If Me.Caption = "" Then Me.Caption = App.ProductName

sHSHead = sck(0).LocalHostName
If sHSHead = "" Then
    sHSHead = "http://" & sck(0).LocalIP
Else
    sHSHead = "http://" & sHSHead
End If

sHSHead = sHSHead & ":" & LTrim$(Str$(sck(0).LocalPort)) + "/"

MZhReaderViaHttpServer.zipProtocolHead = sHSHead

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim hSetting As New CSetting
'hSetting.iniFile = MainFrm.zhtmIni
'With hSetting
'.Save Me, SF_Tag
'.Save Me, SF_CAPTION
'End With
'Set hSetting = Nothing
End Sub

Private Sub Sck_Close(Index As Integer)
    ' disable the timer (so it does not send more data than neccessary)
'    tmrSendData(Index).Enabled = False

    ' make sure the connection is closed
    Do
        sck(Index).Close
        DoEvents
    Loop Until sck(Index).State = sckClosed

End Sub

Private Sub Sck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim K As Integer

    ' look in the control array for a closed connection
    ' note that it's starting to search at index 1 (not index 0)
    ' since index 0 is the one listening on port 80
    Dim lEnd As Long
    lEnd = sck.UBound
    For K = 1 To lEnd
        If sck(K).State = sckClosed Then Exit For
    Next

    ' if all controls are connected, then create a new one
    If K > sck.UBound Then
        K = sck.UBound + 1
        Load sck(K) ' create a new winsock object

'        Load lblFileProgress(K) ' load the label to display the progress on each conection
'        lblFileProgress(K).Top = (lblFileProgress(0).Height + 5) * K
'        lblFileProgress(K).Visible = True
'

'        Load tmrSendData(K) ' load a new timer for the control
''        tmrSendData(K).Enabled = False
''        tmrSendData(K).Interval = 1
    End If

    ' make sure the info structure contains default values (ie: 0's and "")


    ' accept the connection on the closed control or the new control
    sck(K).accept requestID
End Sub

Private Sub Sck_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim rData As String, sHeader As String, RequestedFile As String, ContentType As String

    Dim sFileWanted As String
    Dim fso As New gCFileSystem
    Dim zhUrlNew As zipUrl
    Dim sRealUrl As String

    sck(Index).GetData rData, vbString
    Dim TEMPdATA As String
    TEMPdATA = rData
    rData = TEMPdATA
    If rData Like "GET * HTTP/1.?*" Then
        ' get requested file name
        RequestedFile = LeftRange(rData, "GET ", " HTTP/1.", , ReturnEmptyStr)
        RequestedFile = MZhReaderViaHttpServer.zhServer_DecodeUrl(RequestedFile)
        Debug.Print RequestedFile
        zhUrlNew = zipProtocol_ParseURL(RequestedFile)
                      
        If zhUrlNew.sZipName = "" Or zhUrlNew.sHtmlPath = "" Then
            sFileWanted = fso.BuildPath(App.Path, cHtmlAboutFilename)
        ElseIf isFakeOne(zhUrlNew.sHtmlPath, sRealUrl) Then
            sFileWanted = fso.BuildPath(sTempZip, sRealUrl)
        Else
            sFileWanted = fso.BuildPath(sTempZH, zhUrlNew.sHtmlPath)
           '解压文件
           If fso.PathExists(sFileWanted) = False Then
                MainFrm.myXUnzip zhUrlNew.sZipName, _
                                              zhUrlNew.sHtmlPath, _
                                              sTempZH, _
                                              zhrStatus.sPWD, _
                                              True
            End If
            If fso.PathType(sFileWanted) = LNFolder Then sFileWanted = ""
        End If
'
'        ElseIf zhUrlNew.sZipName <> "" Then
'            sFileWanted = zhUrlNew.sZipName
'        Else
'            sFileWanted = fso.BuildPath(App.Path, cHtmlAboutFilename)
'        End If
'
'        If bFakeHtml And fso.FileExists(sFileWanted) = True Then
'             Dim sHtmlfile As String
'             Dim sTemplateFile As String
'             Dim bUseTemplate As Boolean
'             Dim sTOK As Boolean
'             Dim ftFileWanted As mFile_FileType
'             ftFileWanted = chkFileType(sFileWanted)
'             If ftFileWanted = ftIMG Or ftFileWanted = ftAUDIO Or ftFileWanted = ftVIDEO Then
'             sFileWanted = RequestedFile
'             End If
'              sHtmlfile = fso.BuildPath(sTempzh, fso.GetTempName & ".htm")
'             sTemplateFile = iniGetSetting(, "Viewstyle", "TemplateFile")
'             bUseTemplate = CBoolStr(iniGetSetting(, "ViewStyle", "UseTemplate"))
'             If bUseTemplate Then
'                 sTOK = createHtmlFromTemplate(sFileWanted, sTemplateFile, sHtmlfile)
'                 If sTOK = False Then sTOK = createDefaultHtml(sFileWanted, sHtmlfile)
'                 If sTOK = False Then sHtmlfile = sFileWanted
'             Else
'                 sTOK = createDefaultHtml(sFileWanted, sHtmlfile)
'                 If sTOK = False Then sHtmlfile = sFileWanted
'             End If
'             sFileWanted = sHtmlfile
'        End If
'
        If sFileWanted <> "" And fso.PathExists(sFileWanted) = True Then
        
            Select Case LCase$(fso.GetExtensionName(sFileWanted))
            Case "txt", "text"
                ContentType = "Content-Type: text/plain"
            Case "jpg", "jpeg"
                ContentType = "Content-Type: image/jpeg"
            Case "gif"
                ContentType = "Content-Type: image/gif"
            Case "htm", "html"
                ContentType = "Content-Type: text/html"
            Case "zip"
                ContentType = "Content-Type: application/zip"
            Case "mp3"
                ContentType = "Content-Type: audio/mpeg"
            Case "m3u", "pls", "xpl"
                ContentType = "Content-Type: audio/x-mpegurl"
            Case Else
                ContentType = "Content-Type: */*"
            End Select
            
            ' build the header
            Dim iFile As Integer
            Dim BytesLength As Long
            Dim Buffer() As Byte
            Const BufferLength As Long = 307200
            
            iFile = FreeFile
            Open sFileWanted For Binary Access Read As iFile
    
            sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                    "Server: " & Me.Caption & vbNewLine & _
                    ContentType & vbNewLine & _
                    "Content-Length: " & LOF(iFile) & vbNewLine & _
                     vbNewLine
    
            sck(Index).Tag = sFileWanted
            If sck(Index).State = sckConnected Then
                sck(Index).SendData sHeader
            End If
 
            Do Until EOF(iFile)
                If Loc(iFile) + BufferLength > LOF(iFile) Then
                    BytesLength = LOF(iFile) - Loc(iFile)
                Else
                    BytesLength = BufferLength
                End If
                ReDim Buffer(BytesLength) As Byte
                Get iFile, , Buffer() ' get data from file
                If sck(Index).State = sckConnected Then
                sck(Index).SendData Buffer() ' send thedata on the current connection
                End If
            Loop
            Close iFile
            '删除文件
            'Kill sFileWanted
        Else  ' send "Not Found" if file does not exsist on the share
            If sck(Index).State = sckConnected Then
                sHeader = "HTTP/1.0 404 Not Found" & vbNewLine & "Server: " & Me.Caption & vbNewLine & vbNewLine
                sck(Index).SendData sHeader
            End If
        End If
        
Else    ' sometimes the browser makes "HEAD" requests (but it's not inplemented in this project)
    If sck(Index).State = sckConnected Then
        sHeader = "HTTP/1.0 501 Not Implemented" & vbNewLine & "Server: " & Me.Caption & vbNewLine & vbNewLine
        sck(Index).SendData sHeader
    End If
End If

End Sub
Private Sub sck_SendComplete(Index As Integer)
sck(Index).Close
End Sub

