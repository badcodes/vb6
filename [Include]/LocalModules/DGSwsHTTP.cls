VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDGSwsHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' DGSwsHTTP Class Module by RegX
' Copyright 2002 DGS
' You may freely use this code as long as
' All Copyright information remains intact
'=============================================
'Requires:
'cAppendString.cls
'Microsoft Scripting Runtime
'MSWinsockLib.Winsock (Mswinsck.ocx)
'=============================================
'Description
'Makes downloading HTML files with progress a snap
'Exposes Response Header and FileData seperately

'Usage ---------------------------------------------------------------
'Dim WithEvents wsc As DGSwsHTTP
'Set wsc = New DGSwsHTTP
'wsc.geturl "http://www.microsoft.com/"
'---------------------------------------------------------------------
Option Explicit
Dim cstring As CAppendString
Dim WithEvents ws As MSWinsockLib.Winsock
Attribute ws.VB_VarHelpID = -1
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent ProgressChanged[(arg1, arg2, ... , argn)]
Public Event ProgressChanged(ByVal bytesreceived As Long)
'local variable(s) to hold property value(s)
Public URL As String 'local copy
'local variable(s) to hold property value(s)
Private mvarRemotePort As String 'local copy
Private mvarRemoteHost As String 'local copy
Public totalbytesreceived As Long
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent DownloadComplete[(arg1, arg2, ... , argn)]
Public Event DownloadComplete()
Public FileSize As Variant
Public ResponseHeaderString As String
Private dicResponseHeader As Dictionary
'local variable(s) to hold property value(s)
Public FileData As Variant
Public ResponseCode As String
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent httpError[(arg1, arg2, ... , argn)]
Public Event httpError(errmsg As String, Scode As String)

Public Property Let RemoteHost(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemoteHost = 5
    mvarRemoteHost = vData
End Property

Public Property Get RemoteHost() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemoteHost
    RemoteHost = mvarRemoteHost
End Property

Public Property Let RemotePort(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemotePort = 5
    mvarRemotePort = vData
End Property

Public Property Get RemotePort() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemotePort
    RemotePort = mvarRemotePort
End Property


Public Sub geturl(URL As String)
            'Automatically gets remote host from URL
            If URL & "" = "" Then RaiseEvent httpError("No URL specified", 0)
            cstring.Clear
            Dim RHstart As Long
            Dim RHend As Long
            Dim RH As String
            RHstart = InStr(1, URL, "://", vbTextCompare)
            If RHstart = 0 Then RHstart = 1 Else RHstart = RHstart + 3
            If Len(URL) <= RHstart + 3 Then
                RaiseEvent httpError("Invalid URL", 0)
                Exit Sub
            End If
            RHend = InStr(RHstart + 3, URL, "/", vbTextCompare)
            If RHend = 0 Then RHend = Len(URL) + 1
            RH = Mid(URL, RHstart, RHend - (RHstart))
            
            ' reset winsock
            Me.URL = URL
            ws.Close
            ' connect to server
            ws.RemotePort = Me.RemotePort
            ws.RemoteHost = RH
            ws.Connect
        ' send request

End Sub

Private Sub Class_Initialize()
Set ws = New MSWinsockLib.Winsock
Set cstring = New CAppendString
Set dicResponseHeader = New Dictionary
'defaults
Me.RemotePort = 80
End Sub

Private Sub Class_Terminate()
Set dicResponseHeader = Nothing
Set cstring = Nothing
Set ws = Nothing
End Sub

Private Sub ws_Close()
FileData = cstring.Value
cstring.Clear
RaiseEvent DownloadComplete
End Sub

Private Sub ws_Connect()
    Me.totalbytesreceived = 0
    Me.FileSize = 0
    Me.ResponseHeaderString = ""
    Me.FileData = ""
        Dim Request As String
        ' send request
        'MsgBox Me.URL
        Request = "GET " & Me.URL & " HTTP/1.0" & vbCrLf & vbCrLf
        ws.SendData Request
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
dicResponseHeader.RemoveAll
On Error Resume Next
Dim data As String
Dim arrheader() As String
Dim headkeys() As String
Dim headvals() As String
Dim item As Variant
Dim x As Long
ws.GetData data
        ' get header if this is the begining of the file
        If totalbytesreceived = 0 Then
            ' get server response header
            ResponseHeaderString = Left(data, InStr(1, data, vbCrLf & vbCrLf, vbTextCompare) + 2)
            ' subtract header length and newline from totalbytes received
            totalbytesreceived = bytesTotal - (Len(ResponseHeaderString) + 2)
            cstring.Append Mid(data, Len(ResponseHeaderString) + 2)
            ' see if we have a context length in the header
            
            'parse header data and put in dictionary for easy parsing
            arrheader = Split(ResponseHeaderString, vbCrLf, -1, vbTextCompare)
            For Each item In arrheader
                x = InStr(1, item, " ", vbTextCompare)
                If x > 1 Then
                    dicResponseHeader.Add Left(item, x - 1), Mid(item, x + 1)
                    Debug.Print Left(item, x - 1) & Mid(item, x + 1)
                End If
            Next
                    ResponseCode = Left(dicResponseHeader.Items(0), 3)
                    Me.FileSize = dicResponseHeader.item("Content-Length:")
                    If ResponseCode <> "200" Then
                        RaiseEvent httpError(CStr(dicResponseHeader.Items(0)), ResponseCode)
                    End If
        Else
        cstring.Append data
            totalbytesreceived = totalbytesreceived + bytesTotal
            RaiseEvent ProgressChanged(totalbytesreceived)
        End If
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent httpError(Description, CStr(Scode))
End Sub

