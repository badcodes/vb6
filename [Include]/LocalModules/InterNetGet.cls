VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInterNetGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'欢迎你下载使用本代码，本份代码由程序太平洋提供下载学习之用
'声明:
'1.本站所有代码的版权归原作者所有，如果你使用了在本站下载的源代码
'  引起的一切纠纷(后果)与本站无关,请您尊重原作者的劳动成果！
'2.若本站在代码上有侵权之处请您与站长联系，站长会及时更正。
'下载网站：http://www.daima.com.cn
'论坛：http://www.5ivb.net
'网络实名：代码
'Email:dapha@5ivb.net
'CopyRight 2001-2005 By www.5ivb.net
'整理时间：2004-9-2 19:53:15
Option Explicit
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByRef hInet As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, sBuffer As Byte, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByVal sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const scUserAgent = "Tgwang"
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private mvarUrl As String
Private mvarSaveFile As String
Private mvarConnect As Boolean
Private hOpen As Long, hFile As Long
Private Buffer As String, BufLen As Long
Private RetQueryInfo As Boolean

Public Sub Execute()

    mvarConnect = True

    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)

    If mvarConnect = False Then
        Cancel
        Exit Sub
    End If

    If hOpen = 0 Then
        Cancel
        mvarConnect = False
    Else

        hFile = InternetOpenUrl(hOpen, mvarUrl, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)

        If mvarConnect = False Then
            Cancel
            Exit Sub
        End If

        If hFile = 0 Then
            Cancel
            mvarConnect = False
        Else
            Buffer = Space$(1024)
            BufLen = 1024
            RetQueryInfo = HttpQueryInfo(hFile, 21, Buffer, BufLen, 0)

            If RetQueryInfo Then
                Buffer = Mid$(Buffer, 1, BufLen)
            Else
                Buffer = ""
            End If

        End If

    End If

End Sub

'Public Function FileSize() As Long

'    FileSize = GetHeader("Content-Length")

'End Function

Public Function InternetGet() As Boolean

    Dim sBuffer(1 To 1024) As Byte, Ret As Long
    Dim intfile As Long, LBR As Long
    Dim i As Long

    If mvarConnect = False Then
        Cancel
        InternetGet = False
        Exit Function
    End If

    On Error GoTo OutErr
    Err.Clear
    'If Dir$(mvarSaveFile) > " " Then
    '    Name mvarSaveFile As mvarSaveFile & ".bak"
    'End If
    If Len(Dir$(mvarSaveFile)) > 0 Then
        If MsgBox("目标文件以存在是否覆盖!", vbInformation + vbYesNo, "提示") = vbNo Then
            Cancel
            InternetGet = False
            Exit Function
        End If
    End If
    intfile = FreeFile()

    Open mvarSaveFile For Binary Access Write As #intfile
    Do
        InternetReadFile hFile, sBuffer(1), 1024, Ret
        DoEvents
        If Ret = 1024 Then
            If mvarConnect = False Then
                InternetGet = False
                GoTo Quit
            End If

            Put #1, , sBuffer
        Else
            For i = 1 To Ret
                Put #1, , sBuffer(i)
                DoEvents
            Next i
        End If
        LBR = LBR + Ret
        DoEvents
    Loop Until Ret < 1024
    
    InternetGet = True
    
Quit:
    Close #intfile
    'if Dir$(mvarSaveFile & ".bak") > " " Then
    '   Kill mvarSaveFile
    '  Name mvarSaveFile & ".bak" As mvarSaveFile
    'End If
    Cancel

    Exit Function

OutErr:
    Err.Clear
    Cancel
    Close #intfile
    On Error GoTo 0

End Function

Public Sub Cancel()

    mvarConnect = False
    InternetCloseHandle hOpen
    InternetCloseHandle hFile

End Sub

Public Property Let SaveFile(ByVal FileName As String)

    mvarSaveFile = FileName

End Property

Public Property Let URL(ByVal URL As String)


    mvarUrl = URL

End Property
Public Function GetHeader(Optional hdrName As String) As String

    Dim tmp As Long
    Dim tmp2 As String

    If mvarConnect = False Then
        GetHeader = "0"
        Cancel
        Exit Function
    End If
    If Buffer <> "" Then
        Select Case UCase$(hdrName)
            Case "CONTENT-LENGTH"
                tmp = InStr(Buffer, "Content-Length")
                tmp2 = Mid$(Buffer, tmp + 16, Len(Buffer))
                tmp = InStr(tmp2, Chr$(0))
                GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
            Case "CONTENT-TYPE"
                tmp = InStr(Buffer, "Content-Type")
                tmp2 = Mid$(Buffer, tmp + 14, Len(Buffer))
                tmp = InStr(tmp2, Chr$(0))
                GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
            Case "DATE"
                tmp = InStr(Buffer, "Date")
                tmp2 = Mid$(Buffer, tmp + 6, Len(Buffer))
                tmp = InStr(tmp2, Chr$(0))
                GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
            Case "LAST-MODIFIED"
                tmp = InStr(Buffer, "Last-Modified")
                tmp2 = Mid$(Buffer, tmp + 15, Len(Buffer))
                tmp = InStr(tmp2, Chr$(0))
                GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
            Case "SERVER"
                tmp = InStr(Buffer, "Server")
                tmp2 = Mid$(Buffer, tmp + 8, Len(Buffer))
                tmp = InStr(tmp2, Chr$(0))
                GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
            Case vbNullString
                GetHeader = Buffer
            Case Else
                GetHeader = "0"
        End Select
    Else
        GetHeader = "0"
    End If

End Function

':) 代码格式化插件 V2.3.18 (03-4-22 4:34:02) 24 + 199 = 223 Lines
