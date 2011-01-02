VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const c_urlHandler As String = "pBookurl.exe"
Const c_urlFilter As String = "book://ssreader/*"
Const c_iniFile As String = "config.ini"
Private urlHandler As String
Private urlFilter As String
Public Function wakeUp2(ByVal bookUrl As String, ByVal bookName As String) As Boolean
    On Error Resume Next
    #If DebugMode = 1 Then
        MDebug.DebugFPrint bookUrl
    #End If
    wakeUp2 = False
    If isMyBusiness(bookUrl) = False Then Exit Function
    If Shell(urlHandler & " " & bookUrl & "||" & bookName, vbNormalFocus) = 0 Then Exit Function
    If Err.Number <> 0 Then
        urlHandler = ""
        Exit Function
    End If
    wakeUp2 = True
End Function
Public Function wakeUp3(ByVal bookUrl As String, ByVal bookName As String, ByVal sText As String) As Boolean
    On Error Resume Next
    #If DebugMode = 1 Then
        MDebug.DebugFPrint bookUrl
    #End If
    wakeUp3 = False
    If isMyBusiness(bookUrl) = False Then Exit Function
    If Shell(urlHandler & " " & bookUrl & "||" & bookName & "||" & sText, vbNormalFocus) = 0 Then Exit Function
    If Err.Number <> 0 Then
        urlHandler = ""
        Exit Function
    End If
    wakeUp3 = True
End Function
Public Function wakeUp(ByVal bookUrl As String) As Boolean
    On Error Resume Next
    #If DebugMode = 1 Then
        MDebug.DebugFPrint bookUrl
    #End If
    wakeUp = False
    If isMyBusiness(bookUrl) = False Then Exit Function
    If Shell(urlHandler & " " & bookUrl, vbNormalFocus) = 0 Then Exit Function
    If Err.Number <> 0 Then
        urlHandler = ""
        Exit Function
    End If
    wakeUp = True
End Function
Public Function isMyBusiness(ByVal bookUrl As String) As Boolean
    If urlFilter = "" Or bookUrl = "" Then Exit Function
    'If BookUrl = "" Then Exit Function
    If LCase$(bookUrl) Like LCase$(urlFilter) Then isMyBusiness = True
    'isMyBusiness = startAt(BookUrl, ssReader)
End Function
'Private Function startAt(ByRef strTest As String, ByRef strStart As String, Optional ByRef cmp As VbCompareMethod = vbTextCompare) As Boolean
'    Dim lstrTest As Long
'    Dim lstrStart As Long
'    startAt = False
'    lstrTest = Len(strTest)
'    lstrStart = Len(strStart)
'    If lstrTest < lstrStart Then Exit Function
'    If StrComp(Left$(strTest, lstrStart), strStart, cmp) <> 0 Then Exit Function
'    startAt = True
'End Function

Private Sub Class_Initialize()
    Dim hSet As CLiNInI
    Set hSet = New CLiNInI
    With hSet
        .File = App.Path & "\" & c_iniFile
        urlHandler = .GetSetting("Config", "urlHandler")
        urlFilter = .GetSetting("Config", "urlFilter")
    End With
    
    If urlHandler = "" Then urlHandler = App.Path & "\" & c_urlHandler
    If urlFilter = "" Then urlFilter = c_urlFilter
        
    With hSet
        .File = App.Path & "\" & c_iniFile
        .SaveSetting "Config", "urlHandler", urlHandler
        .SaveSetting "Config", "urlFilter", urlFilter
        .Save
    End With
    
    Set hSet = Nothing
    
    #If DebugMode = 1 Then
        MDebug.DebugFPrint urlHandler
        MDebug.DebugFPrint urlFilter
    #End If
End Sub
