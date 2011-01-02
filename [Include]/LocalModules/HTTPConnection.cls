VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTTPConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, sBuffer As Byte, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByVal sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer


Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const scUserAgent = "Internet Explorer 5.x"
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_NO_UI = 0

Private Const BASE_ERR_NUMBER As Long = "12050"
Private Const CLASS_NAME As String = "CHttpConnection"


Private mHeader As CHttpHeader
Private mHeaderGet As String

Private mvarUrl As String
Private mvarDestination As String
Private mvarConnect As Boolean
Private hOpen As Long, hFile As Long
Private buffer As String, BufLen As Long
Private RetQueryInfo As Boolean

Public Event GetData(Progress As Long)                     '下载进度
'Public Event ErrMassage(Description As String)             '错误信息
Public Event DownLoadOver()
Public Event DataArrived(ByRef Data() As Byte, ByVal Size As Long)
Public Event ProgressChange(ByRef vMin As Long, ByRef vMax As Long)
Public Event StatusChange(ByRef vStatus As String)


''
' Raise errors
Private Sub raiseErr(ByVal lngOffset As Long, ByRef strMsg As String, ByRef strSource As String)
    Err.Raise BASE_ERR_NUMBER + lngOffset, CLASS_NAME & "." & strSource, strMsg
End Sub

Property Let Header(ByVal vHeader As String)
    mHeaderGet = vHeader
End Property

Public Property Get Respone() As CHttpHeader
    Set Respone = mHeader
    
End Property

Public Function Connect() As Boolean
    
    Cancel
    
    Connect = False
    mvarConnect = True

    If hOpen = 0 Then
        Cancel
        RaiseEvent StatusChange("无法创建连接")
        Connect = False
        Exit Function
        'mvarConnect = False
    Else
        
        RaiseEvent StatusChange("打开连接...")
        
        hFile = InternetOpenUrl(hOpen, mvarUrl, mHeaderGet, ByVal Len(mHeaderGet), INTERNET_FLAG_NO_UI, ByVal 0&)

'        If mvarConnect = False Then
'            Cancel
'            Exit Function
'        End If

        If hFile = 0 Then
            Cancel
            RaiseEvent StatusChange("无法连接服务器")
            mvarConnect = False
            Connect = False
            Exit Function
        Else
            RaiseEvent StatusChange("连接成功")
            buffer = Space$(1024)
            BufLen = 1024
            RetQueryInfo = HttpQueryInfo(hFile, 21, buffer, BufLen, 0)
            'Debug.Print Buffer
            If RetQueryInfo Then
                buffer = Mid$(buffer, 1, BufLen)
                mHeader.Init buffer
            Else
                mHeader.Init ""
                buffer = ""
            End If

        End If

    End If

    If InStr(buffer, "HTTP/1.1 200 OK") > 0 Then Connect = True

End Function


Public Sub Retrieve()
Const BufferLength = 512
    
    Dim sBuffer(1 To BufferLength) As Byte, ret As Long
    Dim intfile As Long, LBR As Long
    Dim i As Long
    Dim mTotal As Long
    
    mTotal = mHeader.ConentLength
    
    If mvarConnect = False Then
        Cancel
        Exit Sub
    End If

    On Error GoTo OutErr
    Err.Clear

    'intfile = FreeFile()
    RaiseEvent StatusChange("Start Downloading...")
    'Open mvarDestination For Binary Access Write As #intfile
    Do
        RaiseEvent ProgressChange(LBR, mTotal)
        InternetReadFile hFile, sBuffer(1), BufferLength, ret
        DoEvents
        If ret = 1024 Then
            If mvarConnect = False Then
                GoTo Quit
            End If
            
            RaiseEvent DataArrived(sBuffer, BufferLength)
            
            'Put #intfile, , sBuffer
        ElseIf ret = BufferLength Then
            RaiseEvent DataArrived(sBuffer, BufferLength)
            'Put #intfile, , sBuffer()
        Else
            RaiseEvent DataArrived(sBuffer, ret)
            'For i = 1 To Ret
               
                'Put #intfile, , sBuffer(i)
             ' Next i
        End If
        LBR = LBR + ret
        
        DoEvents
    Loop Until ret < BufferLength
    RaiseEvent ProgressChange(LBR, mTotal)
    RaiseEvent StatusChange("Download Completed")
    RaiseEvent DownLoadOver
    
   
Quit:
   
    Cancel

    Exit Sub

OutErr:
    Err.Clear
    Cancel
  
    'On Error GoTo 0

End Sub

Public Function StartDownload() As Boolean

    Const BufferLength = 512
    
    Dim sBuffer(1 To BufferLength) As Byte, ret As Long
    Dim intfile As Long, LBR As Long
    Dim i As Long
    Dim mTotal As Long
    
    mTotal = mHeader.ConentLength
    
    If mvarConnect = False Then
        Cancel
        StartDownload = False
        Exit Function
    End If

    On Error GoTo OutErr
    Err.Clear

    intfile = FreeFile()
    RaiseEvent StatusChange("正在下载")
    Open mvarDestination For Binary Access Write As #intfile
    Do
        RaiseEvent ProgressChange(LBR, mTotal)
        DoEvents
        InternetReadFile hFile, sBuffer(1), BufferLength, ret
        If ret = 1024 Then
            If mvarConnect = False Then
                StartDownload = False
                GoTo Quit
            End If

            Put #intfile, , sBuffer
        ElseIf ret = BufferLength Then
            Put #intfile, , sBuffer()
        Else
            For i = 1 To ret
                   Put #intfile, , sBuffer(i)
              Next i
        End If
        LBR = LBR + ret
        
        DoEvents
    Loop Until ret < BufferLength
    RaiseEvent ProgressChange(LBR, mTotal)
    RaiseEvent StatusChange("下载完成")
    RaiseEvent DownLoadOver
    
    StartDownload = True
Quit:
    Close #intfile
    'if Dir$(mvarDestination & ".bak") > " " Then
    '   Kill mvarDestination
    '  Name mvarDestination & ".bak" As mvarDestination
    'End If
    Cancel

    Exit Function

OutErr:
    On Error Resume Next
    Err.Clear
    Cancel
    Close #intfile
    raiseErr 3, "文件" & mvarDestination & "正在使用,无法进行操作", "StartDownload"
    'On Error GoTo 0

End Function

Public Sub Cancel()

On Error Resume Next
  
    mvarConnect = False
    InternetCloseHandle hFile
    'InternetCloseHandle hOpen
    
    hFile = 0
    'hOpen = 0
    
    mHeader.Init ""
    buffer = ""

End Sub

Public Property Let Destination(ByVal filename As String)

    mvarDestination = filename

End Property
Public Property Get Destination() As String
    Destination = mvarDestination
End Property
Public Property Let Url(ByVal Url As String)
    mvarUrl = Url
End Property

Public Property Get Url() As String
    Url = mvarUrl
End Property

Private Sub Class_Initialize()
    Set mHeader = New CHttpHeader
     RaiseEvent StatusChange("创建连接...")
     hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
     If hOpen = 0 Then
        RaiseEvent StatusChange("创建连接失败")
        raiseErr 0, "无法创建连接", "Execute"
        mvarConnect = False
    End If
    RaiseEvent StatusChange("成功创建连接")
End Sub

Private Sub Class_Terminate()
    RaiseEvent StatusChange("关闭连接")
    InternetCloseHandle hOpen
    Cancel
    Set mHeader = Nothing
End Sub

