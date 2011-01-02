VERSION 5.00
Begin VB.Form cmdFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BookUrl"
   ClientHeight    =   2148
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5976
   Icon            =   "cmdFrm.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "fgBookUrl|Wakeup"
   MaxButton       =   0   'False
   ScaleHeight     =   2148
   ScaleWidth      =   5976
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmUrl 
      Caption         =   "设置∶"
      Height          =   1068
      Left            =   120
      TabIndex        =   5
      Top             =   132
      Width           =   5724
      Begin VB.TextBox txtBookName 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   840
         TabIndex        =   8
         Top             =   624
         Width           =   4776
      End
      Begin VB.TextBox txtUrl 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   4776
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "URL∶"
         Height          =   264
         Left            =   84
         TabIndex        =   11
         Top             =   264
         Width           =   780
      End
      Begin VB.Label lblBookName 
         Alignment       =   2  'Center
         Caption         =   "书名更正∶"
         Height          =   264
         Left            =   84
         TabIndex        =   9
         Top             =   660
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdReadText 
      Appearance      =   0  'Flat
      Caption         =   "->正文"
      Height          =   320
      Left            =   2220
      TabIndex        =   4
      Top             =   1320
      Width           =   828
   End
   Begin VB.CommandButton cmdReadCata 
      Appearance      =   0  'Flat
      Caption         =   "->目录"
      Height          =   320
      Left            =   1164
      TabIndex        =   3
      Top             =   1320
      Width           =   828
   End
   Begin VB.CommandButton cmdDirectlyDown 
      Appearance      =   0  'Flat
      Caption         =   "->直接下载"
      Default         =   -1  'True
      Height          =   320
      Left            =   4644
      TabIndex        =   2
      Top             =   1296
      Width           =   1140
   End
   Begin VB.CommandButton cmdReadCov 
      Appearance      =   0  'Flat
      Caption         =   "->封面"
      Height          =   320
      Left            =   156
      TabIndex        =   1
      Top             =   1320
      Width           =   828
   End
   Begin VB.CommandButton cmdssReaderDownload 
      Appearance      =   0  'Flat
      Caption         =   "->ssReader"
      Height          =   320
      Left            =   3264
      TabIndex        =   0
      Top             =   1308
      Width           =   1188
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000001&
      Height          =   312
      Left            =   156
      TabIndex        =   10
      Top             =   1752
      Width           =   5676
   End
   Begin VB.Label lblDDE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DDE"
      ForeColor       =   &H80000008&
      Height          =   1968
      Left            =   1812
      TabIndex        =   7
      Top             =   12
      Visible         =   0   'False
      Width           =   816
   End
End
Attribute VB_Name = "cmdFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim urlHander As CbookUrl

Const c_mdbFile As String = "download.mdb"
Const c_Covs = 5
Const c_sCov = "cov"
Const c_Boks = 5
Const c_sBok = "bok"
Const c_Legs = 2
Const c_sLeg = "leg"
Const c_Fows = 20
Const c_sFow = "fow"
Const c_Cats = 30
Const c_sCat = "!00"
Const c_Atts = 30
Const c_sAtt = "att"
Const c_filenameWidth = "6"
Const c_pageWidth = "3"

Dim WithEvents httpHND As CDownLoad
Attribute httpHND.VB_VarHelpID = -1

'Dim WithEvents httpHnd As CDGSwsHTTP



Private Function pageCount(ByRef preUrl As String, ByVal rangeFrom As Long, ByVal rangeTo As Long) As Long
pageCount = rangeTo

'    Dim urlTester As JetCarNetscape
'    Dim curPos As Integer
'    Dim sUrl As String
'
'    Set urlTester = New JetCarNetscape
'    curPos = rangeTo
'    sUrl = buildUrl(preUrl, curPos, c_pageWidth, ".pdg")
'
'    If Not urlTester.IsUrlExist(sUrl) Then
'        curPos = curPos / 2
'        If curPos < rangeFrom Then Exit Function
'
'        sUrl = buildUrl(preUrl, curPos, c_pageWidth, ".pdg")
'        If urlTester.IsUrlExist(sUrl) Then
'            pageCount = pageCount(preUrl, curPos, rangeTo)
'        Else
'            pageCount = pageCount(preUrl, rangeFrom, curPos)
'        End If
'    Else
'        pageCount = curPos
'    End If

End Function
Private Function buildUrl(ByRef preUrl As String, ByRef iPos As Long, ByRef iNumWidth As Long, ByRef sSuffix As String) As String
    buildUrl = preUrl & StrNum(CInt(iPos), CInt(iNumWidth)) & sSuffix
End Function

Private Sub cmdssReaderDownload_Click()
     
    urlHander.setHeader toDownload
    'MsgBox urlHander.toString
    
    ShellExecute Me.hwnd, "open", urlHander.ToString, "", "", 1
    Call downloadLog(urlHander)
    Unload Me
End Sub

Private Sub cmdDirectlyDown_Click()
    Dim strPath As String
    Dim cObject As Control
    
    MWindows.setPosition Me, HWND_BOTTOM
    Load frmDirSelect
    frmDirSelect.Show vbModal, Me
    strPath = frmDirSelect.dirPath
    Unload frmDirSelect
    Set frmDirSelect = Nothing
    
    On Error Resume Next
    For Each cObject In Me.Controls
        If TypeName(cObject) = "CommandButton" Then cObject.Enabled = False
    Next
    Call talkToFlashget(urlHander, lblDDE, strPath)
    For Each cObject In Me.Controls
         If TypeName(cObject) = "CommandButton" Then cObject.Enabled = True
    Next
    
End Sub
'CSEH: ErrMsgBox-xrlin
Private Function talkToFlashget(ByRef urlHnd As mpBookUrl.CbookUrl, ByRef linkDDE As Label, ByVal strSaveIn As String) As Boolean
    '<EhHeader>
    On Error GoTo talkToFlashget_Err
    '</EhHeader>
    Dim sPreUrl As String
    Dim i As Long
'    Dim iFrom As Long
'    Dim iTo As Long
    Dim sSavedIn As String
    Dim sUrl As String
    Dim sFilename As String
    Dim sSavedAS As String

    
    
    
    sSavedIn = strSaveIn & "\" & urlHnd.Bookname & "_" & urlHnd.SS & "\"
    
    xMkdir sSavedIn
    
    Dim sBookinfo As String
    Dim fNum As Integer
    
    sBookinfo = sSavedIn & "bookinfo.dat"
    fNum = FreeFile()
    
    Open sBookinfo For Output As #fNum
    Print #fNum, "[General Information]"
    Print #fNum, "书名=", urlHander.Bookname
    Print #fNum, "作者=", urlHander.Author
    Print #fNum, "页数=", urlHander.Pages
    Print #fNum, "ss号=", urlHander.realSS
    Print #fNum, "下载=", urlHander.URL
    Close #fNum
    
    
    Set httpHND = New CDownLoad
    
    i = 0
    Do While True
        i = i + 1
        sFilename = buildUrl(c_sCov, i, c_pageWidth, ".pdg")
        sUrl = urlHnd.Location & "/" & sFilename
        If Not download(sUrl, sSavedIn & sFilename) Then Exit Do
    Loop
    
    i = 0
    Do While True
        i = i + 1
        sFilename = buildUrl(c_sBok, i, c_pageWidth, ".pdg")
        sUrl = urlHnd.Location & "/" & sFilename
        If Not download(sUrl, sSavedIn & sFilename) Then Exit Do
    Loop
    
    i = 0
    Do While True
        i = i + 1
        sFilename = buildUrl(c_sLeg, i, c_pageWidth, ".pdg")
        sUrl = urlHnd.Location & "/" & sFilename
        If Not download(sUrl, sSavedIn & sFilename) Then Exit Do
    Loop
    
    i = 0
    Do While True
        i = i + 1
        sFilename = buildUrl(c_sFow, i, c_pageWidth, ".pdg")
        sUrl = urlHnd.Location & "/" & sFilename
        If Not download(sUrl, sSavedIn & sFilename) Then Exit Do
    Loop
    
    i = 0
    Do While True
        i = i + 1
        sFilename = buildUrl(c_sCat, i, c_pageWidth, ".pdg")
        sUrl = urlHnd.Location & "/" & sFilename
        If Not download(sUrl, sSavedIn & sFilename) Then Exit Do
    Loop
    
    i = 0
    Do While True
        i = i + 1
        sFilename = buildUrl(c_sAtt, i, c_pageWidth, ".pdg")
        sUrl = urlHnd.Location & "/" & sFilename
        If Not download(sUrl, sSavedIn & sFilename) Then Exit Do
    Loop
    
    sFilename = "bookContent.dat"
    sUrl = urlHnd.Location & "/" & sFilename
    download sUrl, sSavedIn & sFilename
    
    i = 0
    Do While True
        i = i + 1
        sFilename = buildUrl("", i, c_filenameWidth, ".pdg")
        sUrl = urlHnd.Location & "/" & sFilename
        If Not download(sUrl, sSavedIn & sFilename) Then Exit Do
    Loop
    
    
    Set httpHND = Nothing
    
'
'    sPreUrl = urlHnd.Location & "/" & c_sCov
'    iTo = pageCount(sPreUrl, iFrom, c_Covs)
'    For i = iFrom To iTo
'        sUrl = buildUrl(sPreUrl, i, c_pageWidth, ".pdg")
'        AddUrl sUrl, sSavedIn, linkDDE
'    Next
'
'    sPreUrl = urlHnd.Location & "/" & c_sBok
'    iTo = pageCount(sPreUrl, iFrom, c_Boks)
'    For i = iFrom To iTo
'        sUrl = buildUrl(sPreUrl, i, c_pageWidth, ".pdg")
'        AddUrl sUrl, sSavedIn, linkDDE
'    Next
'
'    sPreUrl = urlHnd.Location & "/" & c_sLeg
'    iTo = pageCount(sPreUrl, iFrom, c_Legs)
'    For i = iFrom To iTo
'        sUrl = buildUrl(sPreUrl, i, c_pageWidth, ".pdg")
'        AddUrl sUrl, sSavedIn, linkDDE
'    Next
'
'    sPreUrl = urlHnd.Location & "/" & c_sFow
'    iTo = pageCount(sPreUrl, iFrom, c_Fows)
'    For i = iFrom To iTo
'        sUrl = buildUrl(sPreUrl, i, c_pageWidth, ".pdg")
'        AddUrl sUrl, sSavedIn, linkDDE
'    Next
'
'    sPreUrl = urlHnd.Location & "/" & c_sCat
'    iTo = pageCount(sPreUrl, iFrom, c_Cats)
'    For i = iFrom To iTo
'        sUrl = buildUrl(sPreUrl, i, c_pageWidth, ".pdg")
'        AddUrl sUrl, sSavedIn, linkDDE
'    Next
'
'    sPreUrl = urlHnd.Location & "/" & c_sAtt
'    iTo = pageCount(sPreUrl, iFrom, c_Atts)
'    For i = iFrom To iTo
'        sUrl = buildUrl(sPreUrl, i, c_pageWidth, ".pdg")
'        AddUrl sUrl, sSavedIn, linkDDE
'    Next
'
'    sPreUrl = urlHnd.Location & "/"
'    iTo = pageCount(sPreUrl, iFrom, urlHnd.Pages)
'    For i = iFrom To iTo
'        sUrl = buildUrl(sPreUrl, i, c_filenameWidth, ".pdg")
'        AddUrl sUrl, sSavedIn, linkDDE
'    Next
'
''    Debug.Print pageCount(spreUrl, iFrom, iTo)

    '<EhFooter>
    Exit Function

talkToFlashget_Err:
    MsgBox Err.Description & vbCrLf & _
           "in fgBookUrl.cmdFrm.talkToFlashget "
    '</EhFooter>
End Function
Private Function download(ByRef sUrl As String, ByRef sFileSaved As String) As Boolean
        download = False
        With httpHND
            .URL = sUrl
            .SaveFile = sFileSaved
            If Not .Execute Then Exit Function
            .StartDownLoad
        End With
        download = True
End Function
'Private Sub addPage(ByRef spreUrl As String, ByRef iFrom As Long, ByRef iTo As Long, ByRef iNumWidth As Integer, ByRef sSaveIn As String, ByRef linkDDE As Label)
'    Dim i As Long
'    Dim sUrl As String
'    For i = iFrom To iTo
'        sUrl = buildUrl(spreUrl, i, iNumWidth, ".pdg")
'        AddUrl sUrl, sSavedIn, linkDDE
'    Next
'End Sub
Private Sub AddUrl(ByRef sUrl As String, ByRef sSavedIn As String, ByRef linkDDE As Label)
    Dim httpHND As New CDownLoad
    With httpHND
    .URL = sUrl & "1"
    .SaveFile = "c:\test"
    .Execute
    End With
    

'    With linkDDE
'        .LinkTopic = "FLASHGET|WWW_OPENURL"
'        .LinkItem = sUrl & "," & sSavedIn & "," & sUrl
'        .LinkMode = vbLinkManual
'        .LinkRequest
'    End With
End Sub
'Private Sub cmdRandom_Click()
'
'    Dim L As Integer
'    On Error Resume Next
'    Randomize
'    L = CInt(urlHander.Pages)
'    L = Int(Rnd() * L + 1)
'
'    urlHander.Page = urlHander.page_Text(L)
'    urlHander.setHeader toRead
'
'    ShellExecute Me.hwnd, "open", urlHander.toString, "", "", 1
'    Unload Me
'
'
'End Sub

Private Sub cmdReadCov_Click()
    urlHander.Page = urlHander.page_Cover
    urlHander.setHeader toRead
    ShellExecute Me.hwnd, "open", urlHander.ToString, "", "", 1
    Unload Me
End Sub

Private Sub cmdReadCata_Click()
    urlHander.Page = urlHander.page_Catalog
    urlHander.setHeader toRead
    ShellExecute Me.hwnd, "open", urlHander.ToString, "", "", 1
    Unload Me
End Sub

Private Sub cmdReadText_Click()
    urlHander.Page = urlHander.page_Text(1)
    urlHander.setHeader toRead
    ShellExecute Me.hwnd, "open", urlHander.ToString, "", "", 1
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sUrl As String
    Dim sName As String
    Dim sText As String
    
    Dim sUrlPart() As String
    Dim iU As Long
    Dim iL As Long
    
    On Error Resume Next
    
    Set urlHander = New CbookUrl
    sUrl = Command$
    If Left$(sUrl, 1) = Chr(34) And Right$(sUrl, 1) = Chr(34) Then
        sUrl = Mid$(sUrl, 2, Len(sUrl) - 2)
    End If
    
    sUrlPart = Split(sUrl, "||")
    iU = UBound(sUrlPart)
    iL = LBound(sUrlPart)
    
    If iL > -1 And iU > iL Then
        sUrl = sUrlPart(iL)
        sName = sUrlPart(iL + 1)
        If iU > iL + 1 Then sText = sUrlPart(iL + 1 + 1)
    End If
        
    urlHander.Initialize sUrl
    urlHander.Candownload = "1"
    
    If sText <> "" Then parseText sText, urlHander

    Randomize
    If urlHander.SS = "" Then urlHander.SS = CStr(str2lng(urlHander.Location) * 10 + CLng(urlHander.Pages))
    If urlHander.Author = "" Then urlHander.Author = Environ$("USERNAME")
    If sName = "" Then
        txtBookName.Text = urlHander.Bookname
    Else
        txtBookName.Text = sName
    End If
      
    Dim sClipboard As String
    sClipboard = Clipboard.GetText()
    If InStr(sClipboard, txtBookName.Text) > 0 Then
        txtBookName.Text = sClipboard
        Clipboard.Clear
    End If
        
    urlHander.setHeader toRead
    txtBookName.Text = LTrim$(RTrim$(txtBookName.Text))
    txtUrl.Text = urlHander.ToString
    MWindows.setPosition Me, HWND_TOPMOST
End Sub

Private Sub Form_Terminate()
    Set urlHander = Nothing
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set urlHander = Nothing
End Sub

Public Sub Connect(ByVal sUrl As String)
    urlHander.Initialize sUrl
    urlHander.Candownload = "1"
    txtBookName.Text = urlHander.Bookname
End Sub

'Private Sub httpHnd_DownloadComplete()
'
'End Sub
'
'Private Sub httpHnd_httpError(errmsg As String, Scode As String)
'    Err.Raise 19811205, Scode, errmsg
'End Sub
Private Sub httpHND_GetData(Progress As Long)
    Dim sMsg As String
    With httpHND
        sMsg = "[" & Format(Progress / CLng(.GetHeader("Content-Length")), "0.0%") & "]" & .SaveFile
    End With
    displayProgress sMsg
End Sub

'Private Sub httpHnd_ProgressChanged(ByVal bytesreceived As Long)

'End Sub

Private Sub displayProgress(sMsg As String)
    lblStatus.Caption = sMsg
End Sub

Private Sub txtBookName_Change()
    urlHander.Bookname = txtBookName.Text
End Sub

Private Function str2lng(ByRef str As String) As Long
    Dim i As Long
    Dim iEnd As Long
    Dim c As String
    Dim r As Long
    
    r = 0
    iEnd = Len(str)
    For i = 1 To iEnd
        c = Mid$(str, i, 1)
        r = r + AscW(c)
    Next
    If r < 0 Then r = 0 - r
    str2lng = r
    
End Function

Public Sub downloadLog(ByRef bUrl As CbookUrl)
    Dim sMDB As String
    Dim dbase As Database
    Dim tdef As TableDef
    Dim rc As Recordset
    
    sMDB = App.Path & "\" & c_mdbFile
    
    Set dbase = openDatabase(sMDB)
     
    If dbase Is Nothing Then
        Set dbase = newDatabase(sMDB)
        If dbase Is Nothing Then Exit Sub
        Set tdef = newTable(dbase, "Log")
        addField tdef, "ssid"
        addField tdef, "bookname"
        addField tdef, "author"
        addField tdef, "pages"
        addField tdef, "date"
        addField tdef, "url"
        addField tdef, "download"
        addTable dbase, tdef
        Set tdef = Nothing
    End If
    
    If dbase Is Nothing Then Exit Sub
    
    Set tdef = getTable(dbase, "Log")
    Set rc = getRecord(tdef)
    
    rc.AddNew
    With rc
        .Fields("download").Value = Date$ & " " & Time$
        .Fields("ssid").Value = bUrl.realSS
        .Fields("bookname").Value = bUrl.Bookname
        .Fields("author").Value = bUrl.Author
        .Fields("pages").Value = bUrl.Pages
        .Fields("date").Value = bUrl.getParam("date")
        .Fields("url").Value = bUrl.URL
    End With
    
    rc.Update
        
    Set rc = Nothing
    Set tdef = Nothing
    Set dbase = Nothing
    
End Sub


Private Sub parseText(ByVal sText As String, ByRef bUrl As CbookUrl)
' "※ 五代史通俗演义 发表评论  添加个人书签  点击次数：0  请双击书名浏览
'作者：蔡东藩　  出版社： 出版日期：1981年2月第1版  页数：584"
Dim sAuthor As String
Dim sDate As String
Dim sPage As String

sAuthor = RTrim$(LTrim$(LeftRange(sText, "作者：", "出版社", ReturnEmptyStr)))
sDate = RTrim$(LTrim$(LeftRange(sText, "出版日期：", "页数：", ReturnEmptyStr)))
sPage = RTrim$(LTrim$(LeftRight(sText, "页数：", ReturnEmptyStr)))

If sAuthor <> "" Then bUrl.Author = sAuthor
If sDate <> "" Then bUrl.setParam "date", sDate
If sPage <> "" Then bUrl.Pages = sPage

'MsgBox sAuthor & vbCrLf & sDate & vbCrLf & sPage

End Sub
