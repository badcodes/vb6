VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "SysTray.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "zhServer"
   ClientHeight    =   3096
   ClientLeft      =   120
   ClientTop       =   684
   ClientWidth     =   4680
   Icon            =   "frmServer.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "ExecCmd"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3096
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtDDE 
      Height          =   300
      Left            =   1020
      LinkTopic       =   "LBlueSky|ExecCmd"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2055
      Width           =   1215
   End
   Begin SysTrayCtl.cSysTray cSysTray 
      Left            =   2565
      Top             =   900
      _ExtentX        =   910
      _ExtentY        =   910
      InTray          =   0   'False
      TrayIcon        =   "frmServer.frx":2CFA
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin MSWinsockLib.Winsock sck 
      Index           =   0
      Left            =   660
      Top             =   450
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Image imgRequesting 
      Enabled         =   0   'False
      Height          =   192
      Left            =   2232
      Picture         =   "frmServer.frx":2ED4
      Top             =   228
      Width           =   192
   End
   Begin VB.Image imgSendingData 
      Enabled         =   0   'False
      Height          =   192
      Left            =   2928
      Picture         =   "frmServer.frx":325E
      Top             =   216
      Width           =   192
   End
   Begin VB.Image imgNothingTodo 
      Enabled         =   0   'False
      Height          =   192
      Left            =   1656
      Picture         =   "frmServer.frx":35E8
      Top             =   216
      Width           =   192
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "System"
      Begin VB.Menu mnuSystemCstart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuSystemCstop 
         Caption         =   "S&top"
      End
      Begin VB.Menu mnuSet11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemCconfig 
         Caption         =   "&Config"
      End
      Begin VB.Menu fdsfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemCshutDown 
         Caption         =   "Shut&Down"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hssZh As HttpServerSet
Public iniHsszh As String
Dim sHSHead As String
Dim sTempFolder As String
Enum hsErrorCome
    hecErrServerAlreadyRun = 1
End Enum
Private Const sImgNothingtoDo = "images\NothingTodo.ico"
Private Const sImgRequesting = "images\Requesting.ico"
Private Const sImgSendingData = "images\SendingData.ico"
Private Const conErrLogFile = "err.log"
Private sErrLogFile As String
Private WithEvents lUnzip As cUnzip
Attribute lUnzip.VB_VarHelpID = -1
Private PASSWORD As String
Private InvaildPassword As Boolean

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
Cancel = False
If CmdStr <> "" Then ShellExecute Me.hwnd, "open", CmdStr, "", "", 0
End Sub

'CSEH: ErrResumeNext
Private Sub Form_Load()
Dim ret As Long
Dim fso As New FileSystemObject
Dim sImgPath As String
Dim cmdLine As String

If App.PrevInstance = True Then
    Me.LinkTopic = ""          ' 这两行用于清除新运行的程序的DDE服务器属性，
    Me.LinkMode = 0
End If


sTempFolder = fso.GetSpecialFolder(TemporaryFolder)

cmdLine = LeftDelete(RightDelete(Command$, Chr(34)), Chr(34))
startServer cmdLine
    
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

sErrLogFile = fso.BuildPath(App.Path, conErrLogFile)
sImgPath = fso.BuildPath((App.Path), sImgNothingtoDo)
If fso.FileExists(sImgPath) Then imgNothingTodo.Picture = LoadPicture(sImgPath)
sImgPath = fso.BuildPath((App.Path), sImgRequesting)
If fso.FileExists(sImgPath) Then imgRequesting.Picture = LoadPicture(sImgPath)
sImgPath = fso.BuildPath((App.Path), sImgSendingData)
If fso.FileExists(sImgPath) Then imgSendingData.Picture = LoadPicture(sImgPath)
            
Set cSysTray.TrayIcon = imgNothingTodo.Picture


'记录原来的window程序地址
'preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
'用自定义程序代替原来的window程序
'ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf wndproc)

With cSysTray
   .InTray = True
   .TrayTip = MApp.aboutApp & vbCrLf & "Servering " & sck(0).LocalIP & ":" & sck(0).LocalPort
End With


'modZhServer.sHttpServer = sHSHead

End Sub


Private Sub lUnzip_PasswordRequest(sPassword As String, ByVal sName As String, bCancel As Boolean)

bCancel = False
Static lastName As String

If InvaildPassword = False And PASSWORD <> "" Then
    sPassword = PASSWORD
    If sName = lastName Then
        InvaildPassword = True
    Else
        lastName = sName
    End If
Else
    sPassword = InputBox(lUnzip.ZipFile & vbCrLf & sName & " Request For Password", "Password", "")
    If sPassword <> "" Then
        InvaildPassword = False
        PASSWORD = sPassword
    Else
    bCancel = True
    End If
End If
    
End Sub

Private Sub mnuSystem_Click()
If sck(0).State = sckClosed Then
    mnuSystemCstart.Enabled = True
    mnuSystemCstop.Enabled = False
Else
    mnuSystemCstart.Enabled = False
    mnuSystemCstop.Enabled = True
End If
End Sub

Private Sub mnuSystemCconfig_Click()
Load frmOptions
frmOptions.Show
End Sub

Private Sub mnuSystemCshutDown_Click()
    shutDownServer
End Sub

Private Sub mnuSystemCstart_Click()
    startServer ""
End Sub

Private Sub mnuSystemCstop_Click()
    stopServer
End Sub

Private Sub Sck_Close(Index As Integer)

    Set cSysTray.TrayIcon = imgNothingTodo.Picture
    ' disable the timer (so it does not send more data than neccessary)
'    tmrSendData(Index).Enabled = False

    ' make sure the connection is closed
    Do
        sck(Index).Close
        DoEvents
    Loop Until sck(Index).State = sckClosed


End Sub

Private Sub Sck_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    Set cSysTray.TrayIcon = imgRequesting.Picture
    Dim K As Integer

    ' look in the control array for a closed connection
    ' note that it's starting to search at index 1 (not index 0)
    ' since index 0 is the one listening on port 80
    Dim lEnd As Long
    lEnd = sck.UBound
    For K = 1 To lEnd
        If sck(K).State = sckClosed Then Exit For
    Next K

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
    sck(K).Accept requestID
End Sub

Private Sub Sck_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim rData As String, sHeader As String, RequestedFile As String, ContentType As String
    Dim sFilewanted As String
    Dim fso As New FileSystemObject
    Dim hsTempFile As New CTempFile
    Dim sErrMsg As String

    Set cSysTray.TrayIcon = imgSendingData.Picture
    
    sck(Index).GetData rData, vbString

    If rData Like "GET * HTTP/1.?*" Then
        ' get requested file name
        RequestedFile = LeftRange(rData, "GET ", " HTTP/1.", , ReturnEmptyStr)
        RequestedFile = modHttpServer.hs_DecodeUrl(RequestedFile)
        
        Dim zhUrlNew As HSUrl
        Dim unzipResult As unzReturnCode
       ' RequestedFile = DecodeUrl(RequestedFile, CP_UTF8) 'from Modstring

        zhUrlNew = modHttpServer.hs_ParseUrl(RequestedFile)

        Select Case zhUrlNew.urlType
        Case zipHSUrl '兼有 zhurlnew.mainPart zhurlnew.secondPart

                If zhUrlNew.fileType = hsFTElse Or zhUrlNew.IsFakeHtml = False Then
                   unzipResult = infoUnzip(zhUrlNew.mainPart, zhUrlNew.secondPart, sTempFolder, False, "")
                   If unzipResult = PK_OK Then
                       sFilewanted = fso.BuildPath(sTempFolder, fso.GetFileName(zhUrlNew.secondPart))
                       hsTempFile.add sFilewanted
                   Else
                       sFilewanted = ""
                   End If
                 Else
                   sFilewanted = RequestedFile
                 End If
                          
        Case fileHSUrl '文件
             If zhUrlNew.fileType = hsFTElse Then
                sFilewanted = zhUrlNew.mainPart
              Else
                sFilewanted = RequestedFile
              End If
        Case folderHSUrl '文件夹
             zhUrlNew.IsFakeHtml = True
             sFilewanted = fso.GetBaseName(zhUrlNew.mainPart)
             If sFilewanted = "" Then sFilewanted = fso.GetDrive(zhUrlNew.mainPart).DriveLetter & "："
             sFilewanted = fso.BuildPath(sTempFolder, sFilewanted & ".txt")
             hsTempFile.add sFilewanted
             If hs_CreateIndex(zhUrlNew.mainPart, sHSHead, sFilewanted) = False Then
                sErrMsg = sErrMsg & vbCrLf & "Can't Create File List of " & zhUrlNew.mainPart
                sFilewanted = ""
             Else
             hsTempFile.add sFilewanted
             End If
        Case nullHSUrl
             sFilewanted = ""
        End Select
        
        If sFilewanted = "" Or fso.FileExists(sFilewanted) = False Then
            sErrMsg = sErrMsg & "File Not exist >" & sFilewanted
            sFilewanted = logErrorFile(sErrMsg)
            zhUrlNew.IsFakeHtml = True
        End If
        
        
createHtmlTemplate:
        
        If zhUrlNew.IsFakeHtml And sFilewanted <> "" Then
            Dim sHtmlfile As String
            Dim sTemplateFile As String
            Dim bUseTemplate As Boolean
            Dim sTOK As Boolean
            sHtmlfile = fso.BuildPath(sTempFolder, fso.GetTempName & ".htm")
            hsTempFile.add sHtmlfile
            sTemplateFile = hssZh.sTemplateFile
            bUseTemplate = hssZh.bUseTemplate
            If bUseTemplate Then
                sTOK = modHttpServer.hs_createHtmlFromTemplate(sFilewanted, sTemplateFile, sHtmlfile)
                If sTOK = False Then sTOK = modHttpServer.hs_createDefaultHtml(sFilewanted, sHtmlfile)
                If sTOK = False Then
                    sHtmlfile = sFilewanted
                    sErrMsg = sErrMsg & vbCrLf & "Can't Use Template."
                End If
            Else
                sTOK = modHttpServer.hs_createDefaultHtml(sFilewanted, sHtmlfile)
                If sTOK = False Then sHtmlfile = sFilewanted
            End If
            sFilewanted = sHtmlfile
        End If
        
        If sFilewanted = "" Or fso.FileExists(sFilewanted) = False Then
            sFilewanted = logErrorFile(sErrMsg)
        End If


        If fso.FileExists(sFilewanted) = True Then

            Select Case LCase$(fso.GetExtensionName(sFilewanted))
            Case "txt", "text"
                ContentType = "Content-Type: text/plain"
            Case "jpg", "jpeg"
                ContentType = "Content-Type: image/jpeg"
            Case "gif"
                ContentType = "Content-Type: image/gif"
            Case "htm", "html"
                ContentType = "Content-Type: text/html"
            Case "zip", "zhtm"
                sFilewanted = getZipFirstFile(sFilewanted, sHSHead, sZipUrlSep)
                If chkFileType(sFilewanted) = ftIE Then
                    zhUrlNew.IsFakeHtml = False
                Else
                    zhUrlNew.IsFakeHtml = True
                End If
                hsTempFile.add sFilewanted
                GoTo createHtmlTemplate
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
            Open sFilewanted For Binary Access Read As iFile
    
            sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                    "Server: " & hssZh.sName & vbNewLine & _
                    ContentType & vbNewLine & _
                    "Content-Length: " & LOF(iFile) & vbNewLine & _
                     vbNewLine
    
            sck(Index).Tag = sFilewanted
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
            If hs_isTempFile(sFilewanted) Then hsTempFile.add sFilewanted
        Else  ' send "Not Found" if file does not exsist on the share
            If sck(Index).State = sckConnected Then
                sHeader = "HTTP/1.0 404 Not Found" & vbNewLine & "Server: " & hssZh.sName & vbNewLine & vbNewLine
                sck(Index).SendData sHeader
            End If
        End If
        
Else    ' sometimes the browser makes "HEAD" requests (but it's not inplemented in this project)
    If sck(Index).State = sckConnected Then
        sHeader = "HTTP/1.0 501 Not Implemented" & vbNewLine & "Server: " & hssZh.sName & vbNewLine & vbNewLine
        sck(Index).SendData sHeader
    End If
End If

End Sub
Private Sub sck_SendComplete(Index As Integer)
Sck_Close Index
End Sub



Private Sub cSysTray_MouseDown(Button As Integer, id As Long)

If Button = 2 Then
    Dim cpos As POINTAPI
    GetCursorPos cpos
    TrackPopupMenu GetSubMenu(GetMenu(Me.hwnd), 0), (TPM_LEFTALIGN Or TPM_RIGHTBUTTON), cpos.x, cpos.y, 0, Me.hwnd, vbNull
End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim ret As Long
     '取消Message的截取，使之送往原来的window程序
     'ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, preWinProc)

End Sub

Private Sub startServer(ByVal cmdLine As String)



If App.PrevInstance = True And cmdLine = "" Then
    Unload Me
    End
    Exit Sub
End If

iniHsszh = bddir(App.Path) & App.ProductName & ".ini"
modHttpServer.hs_getServerSetting iniHsszh, hssZh

If App.PrevInstance = True Then
    Dim DDEmsg As String
    sHSHead = hssZh.sHostName
    If sHSHead = "" Then sHSHead = hssZh.sIP
    If sHSHead = "" Then sHSHead = "127.0.0.1"
    sHSHead = "http://" & sHSHead
    sHSHead = sHSHead & ":" & LTrim(Str$(hssZh.sPort)) + "/"
    DDEmsg = sHSHead & cmdLine
    LinkAndSendMessage DDEmsg
    Unload Me
    End
    Exit Sub
End If

sck(0).LocalPort = CLngStr(hssZh.sPort)

sck(0).Listen
DoEvents
hssZh.sPort = CStr(sck(0).LocalPort)
hssZh.sIP = sck(0).LocalIP
hssZh.sHostName = sck(0).LocalHostName
With hssZh
If .sName = "" Then .sName = App.ProductName
.sVersion = "Build" & Str$(App.Major) & "." & Str$(App.Minor) & "." & Str$(App.Revision)
End With

modHttpServer.hs_saveServerSetting iniHsszh, hssZh

sHSHead = sck(0).LocalHostName
If sHSHead = "" Then
    sHSHead = "http://" & sck(0).LocalIP
Else
    sHSHead = "http://" & sHSHead
End If

sHSHead = sHSHead & ":" & LTrim(Str$(sck(0).LocalPort)) + "/"

'If cmdline <> "" Then cmdline = OpenZipMain(cmdline, sHSHead, sZipUrlSep)
If cmdLine <> "" Then ShellExecute Me.hwnd, "open", sHSHead & cmdLine & sZipUrlSep & "/", "", "", 0

End Sub

Private Function sErrInfo(hsec As hsErrorCome) As String
    Select Case hsec
        Case hecErrServerAlreadyRun
            sErrInfo = App.ProductName & "'s been Runing Already!"
    End Select
End Function

Private Sub shutDownServer()
    stopServer
    Unload Me
End Sub

Private Sub stopServer()
    Dim lEnd As Long
    Dim l As Long
    lEnd = sck.count - 1
    For l = lEnd To 1 Step -1
        sck(l).Close
        Unload sck(l)
    Next
    sck(0).Close
End Sub

Public Function HtmlAboutFile(sFilename) As String

    Dim fso As New Scripting.FileSystemObject
    Dim fsoTS As Scripting.TextStream


        Set fsoTS = fso.CreateTextFile(sFilename, True)
        With fsoTS
            .WriteLine "<html>"
            .WriteLine "<head>"
            .WriteLine "<Title>A Flying Http File Server</title>"
            .WriteLine "<meta http-equiv=Content-Type content=" & Chr$(34) & "text/html; charset=us-ascii" & Chr$(34) & ">"
            .WriteLine "</head>"
            .WriteLine "<body  background=images\bg.jpg >"
            .WriteLine "<p align=right ><span lang=EN-US style='font-size:24.0pt;font-family:TAHOMA,Courier New'>" & App.ProductName & " (Build" & Str$(App.Major) + "." + Str$(App.Minor) & "." & Str$(App.Revision) & ")</span></p>"
            .WriteLine "<p align=right ><span lang=EN-US style='font-size:24.0pt;font-family:TAHOMA,Courier New'>" & App.LegalCopyright & "</span></span></p>"
            .WriteLine "</body>"
            .WriteLine "</html>"
        End With
        fsoTS.Close

   HtmlAboutFile = sFilename
   
End Function

Public Function logErrorFile(ByVal sErrMsg As String) As String

    Dim fso As New Scripting.FileSystemObject
    Dim fsoTS As Scripting.TextStream


        Set fsoTS = fso.CreateTextFile(sErrLogFile, True)
        With fsoTS
            .WriteLine "---------------------------------------------------------"
            .WriteLine sErrMsg
            .WriteLine "---------------------------------------------------------"
            .WriteLine App.ProductName & " (Build" & Str$(App.Major) + "." + Str$(App.Minor) & "." & Str$(App.Revision) & ")"
            .WriteLine App.LegalCopyright
            .WriteLine Date$
        End With
        fsoTS.Close
        logErrorFile = sErrLogFile
End Function

Private Sub LinkAndSendMessage(ByVal Msg As String)
'picDDE.LinkMode = 0               '--
'picDDE.LinkTopic = "P1|FormDDE"   '  |______连接DDE程序并发送数据/参数
'picDDE.LinkMode = 2               '  |      “|”为管道符，是“退格键”旁边的竖线，
'picDDE.LinkExecute Msg            '--        不是字母或数字！
'
't = picDDE.LinkTimeout     '--
'picDDE.LinkTimeout = 1     '  |______终止DDE通道。当然，也可以用别的方法
'picDDE.LinkMode = 0        '  |      这里用的是超时强制终止的方法
'picDDE.LinkTimeout = t     '--
Dim t As Long
txtDDE.LinkMode = 0
txtDDE.LinkTopic = "LBlueSky|ExecCmd"
txtDDE.LinkMode = 2
txtDDE.LinkExecute Msg
t = txtDDE.LinkTimeout
txtDDE.LinkTimeout = 1
txtDDE.LinkMode = 0
txtDDE.LinkTimeout = t

End Sub

Public Function infoUnzip(sZipName As String, FileToUnzip As String, UnzipTo As String, bPreserverPath As Boolean, Optional sPWD As String) As unzReturnCode
    If sPWD <> "" Then
        PASSWORD = sPWD
        InvaildPassword = False
    End If
    Set lUnzip = New cUnzip
    With lUnzip
        .AddFileToPreocess toUnixPath(FileToUnzip)
        .ZipFile = toUnixPath(sZipName)
        .UnzipFolder = toUnixPath(UnzipTo)
        .UseFolderNames = bPreserverPath
    End With
    infoUnzip = lUnzip.unzip
    Set lUnzip = Nothing
End Function
Public Function getCommentText(sZip As String, Optional sPWD As String)
        Set lUnzip = New cUnzip
        If sPWD <> "" Then
            PASSWORD = sPWD
            InvaildPassword = False
        End If
        lUnzip.ZipFile = sZip
        getCommentText = lUnzip.GetComment
        Set lUnzip = Nothing
End Function

Public Function getZipFirstFile(ByVal thisfile As String, ByVal sHttpHead As String, ByVal sUrlSep As String) As String

    Dim fso As New Scripting.FileSystemObject
    Dim firstfile As String
    Dim unzipTemp As String
    Dim sTmpfile As String
    Dim ts As Scripting.TextStream
    Dim sHref As String
    

    
    If fso.FileExists(thisfile) = False Then Exit Function
    
    unzipTemp = Environ$("temp")
    Dim sTmpText As String
 
    sTmpText = getCommentText(thisfile)

 
    firstfile = LeftRange(sTmpText, "defaultfile", vbCrLf, vbTextCompare, ReturnEmptyStr)
    If firstfile = "" Then firstfile = LeftRange(sTmpText, "defaultfile", vbLf, vbTextCompare, ReturnEmptyStr)
    firstfile = LeftRight(firstfile, "=", vbTextCompare, ReturnEmptyStr)
    firstfile = Trim(firstfile)
    
    If firstfile = "" Then
        
        Set lUnzip = New cUnzip
        Dim uzs As New CZipItems
        'Dim uzi As CZipItem
        lUnzip.ZipFile = thisfile
        lUnzip.getZipItems uzs
        
        Dim sZipFiles() As String
        Dim lzipFilescount As Long
        Dim sArrHtmfile() As String
        Dim lHtmFileCount As Long
        Dim sExtName As String
        Dim sDefaultfile As String
        Dim lEnd As Long
        Dim m As Long
        
        lEnd = uzs.count
        For m = 1 To lEnd
        If Right(uzs(m).FileName, 1) <> "/" Then
            ReDim Preserve sZipFiles(lzipFilescount) As String
            sZipFiles(lzipFilescount) = uzs(m).FileName
            lzipFilescount = lzipFilescount + 1
        End If
        Next
        Set uzs = Nothing
        Set lUnzip = Nothing
        
        lEnd = lzipFilescount - 1
        For m = 0 To lEnd
            sExtName = LCase$(fso.GetExtensionName(sZipFiles(m)))
            
        '    If sExtName = LCase$(cTxtIndex) Then
        '        sDefaultfile = sZipFiles(m)
        '        loadCmdLine = starthttp(thisfile, sHttpServerHead, sDefaultfile)
        '        Exit Function
        '    End If
            
            If sExtName = "htm" Or sExtName = "html" Then
                If IsWebsiteDefaultFile(sZipFiles(m)) Then
                    ReDim Preserve sArrHtmfile(lHtmFileCount) As String
                    sArrHtmfile(lHtmFileCount) = sZipFiles(m)
                    lHtmFileCount = lHtmFileCount + 1
                End If
            End If
        Next
        
        If lHtmFileCount > 1 Then
            sDefaultfile = sArrHtmfile(0)
            QuickSortFiles sArrHtmfile, 0, lHtmFileCount - 1
            firstfile = findDefaultHtml(sArrHtmfile)
        ElseIf lHtmFileCount = 1 Then
            firstfile = sArrHtmfile(0)
        End If

    End If

    If firstfile <> "" Then
        sTmpfile = fso.BuildPath(unzipTemp, fso.GetTempName & ".htm")
        sHref = sHttpHead & thisfile & sUrlSep & "/"
        sHref = toUnixPath(sHref)
        firstfile = toUnixPath(firstfile)
        firstfile = LeftDelete(firstfile, "/")
        sHref = sHref & firstfile
        If chkFileType(sHref) <> ftIE Then sHref = sHref & sFakeHtmlTrail
        Dim shtmlContent As String
        Set ts = fso.CreateTextFile(sTmpfile, True)
        shtmlContent = "<script language=!Javascript!>" & vbCrLf
        shtmlContent = shtmlContent & "document.location=!" & sHref & "!;" & vbCrLf
        shtmlContent = shtmlContent & "</script>"
        shtmlContent = Replace(shtmlContent, "!", Chr$(34))
        ts.Write shtmlContent
        getZipFirstFile = sTmpfile
    
    End If

    If getZipFirstFile = "" Then
         Dim fReal As String
         Dim tsContent As String
         Dim sIndexName As String
         sIndexName = fso.GetBaseName(thisfile)
         If sIndexName = "" Then Exit Function
         fReal = fso.BuildPath(unzipTemp, sIndexName)
         fReal = hs_getTempFileName(fReal)
         If fso.FileExists(fReal) Then fso.DeleteFile fReal, True
         Set ts = fso.OpenTextFile(fReal, ForWriting, True)
         tsContent = "<table width=!100%! border=0 >"
         tsContent = tsContent & "<tr><td align=!center!>"
         tsContent = tsContent & "<table><tr><td style=!line-height: 150%!>"
         lEnd = lzipFilescount - 1
         For m = 0 To lEnd
         sHref = sHttpHead & thisfile & sUrlSep & "/" & sZipFiles(m)
         sHref = toUnixPath(sHref)
         If chkFileType(sHref) <> ftIE Then sHref = sHref & sFakeHtmlTrail
             tsContent = tsContent & "&gt;&gt;&nbsp;<a href=!" & _
                 sHref & "!>" & fso.GetBaseName(sZipFiles(m)) & "</a>" & vbCrLf
         Next
         tsContent = tsContent & "</td></tr></table></td></tr></table>"
         tsContent = Replace(tsContent, "!", Chr$(34))
         ts.Write tsContent
         ts.Close
         getZipFirstFile = fReal
    End If
End Function

