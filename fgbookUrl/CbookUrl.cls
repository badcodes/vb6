VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CbookUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colParam As Collection
Private colValue As Collection

Private m_Head As String ' "book://ssreader/"
Private m_ssid As String

Public Enum headerType
    toRead = 1
    toDownload = 2
End Enum

Public Property Get Page() As String
    Dim sUrl As String
    sUrl = getParam("url")
    Page = RightRight(sUrl, "/", vbBinaryCompare, ReturnEmptyStr)
    If Page = "" Then Page = RightRight(sUrl, "\", vbBinaryCompare, ReturnEmptyStr)
End Property

Public Property Let Page(ByVal Value As String)
    setParam "url", Location & "/" & Value
End Property

Public Property Get Location() As String
    On Error Resume Next
    Dim sUrl As String
    sUrl = getParam("url")
    Location = RightLeft(sUrl, "/", vbBinaryCompare, ReturnEmptyStr)
    If Location = "" Then Location = RightLeft(sUrl, "\", vbBinaryCompare, ReturnEmptyStr)
End Property

Public Property Let Location(ByVal Value As String)
    setParam "url", Page & "/" & Value
End Property
Public Property Get realSS() As String
    realSS = m_ssid
End Property
Public Property Get SS() As String
    SS = getParam("ss")
End Property

Public Property Let SS(ByVal Value As String)
    setParam "ss", Value
End Property
Public Property Get Bookname() As String
    Bookname = getParam("bookname")
End Property

Public Property Let Bookname(ByVal Value As String)
    setParam "bookname", Value
End Property
Public Property Get Pages() As String
    Pages = getParam("pages")
End Property

Public Property Let Pages(ByVal Value As String)
    setParam "pages", Value
End Property

Public Property Get Candownload() As String
    Candownload = getParam("candownload")
End Property

Public Property Let Candownload(ByVal Value As String)
    setParam "candownload", Value
End Property

Public Property Get Author() As String
    Author = getParam("author")
End Property

Public Property Let Author(ByVal Value As String)
    setParam "author", Value
End Property
Public Property Get Url() As String
    Url = getParam("url")
End Property

Public Property Let Url(ByVal Value As String)
    setParam "url", Value
End Property


Public Sub Initialize(ByVal sUrl As String)
'book://ssreader/e0?url=http://XXX/!00001.pdg&&&&&pages=264&bookname=Â³Ñ¸ÊÖ¸åÈ«¼¯

    Dim cTest As CHandler


    Dim vParam() As String
    Dim iLow  As Long
    Dim iUp As Long
    Dim i As Long
    
    Dim sParam As String
    Dim sValue As String
    
    
    Set cTest = New CHandler
    If cTest.isMyBusiness(sUrl) = False Then Exit Sub
    Set cTest = Nothing
    On Error Resume Next
    
    m_Head = LeftLeft(sUrl, "?", vbBinaryCompare, ReturnEmptyStr)
    sUrl = LeftRight(sUrl, "?", vbBinaryCompare, ReturnEmptyStr)
    
    
    vParam = Split(sUrl, "&")
    iLow = LBound(vParam)
    iUp = UBound(vParam)
    
    Set colParam = New Collection
    
    For i = iLow To iUp
        sParam = LeftLeft(vParam(i), "=", vbBinaryCompare, ReturnEmptyStr)
        sValue = LeftRight(vParam(i), "=", vbBinaryCompare, ReturnEmptyStr)
        If sParam <> "" Then
            colParam.Remove sParam
            colParam.Add sParam, sParam
            colValue.Add sValue, sParam
        End If
    Next

    m_ssid = getParam("ss")

End Sub

Public Function getParam(ByRef sParam As String) As String
    On Error Resume Next
    getParam = colValue.Item(sParam)
End Function

Public Sub setParam(ByRef sParam As String, ByRef sValue As String)
    On Error Resume Next

    
    colParam.Remove sParam
    colValue.Remove sParam
    
    colParam.Add sParam, sParam
    colValue.Add sValue, sParam

    
End Sub

Private Sub Class_Initialize()
    Set colParam = New Collection
    Set colValue = New Collection
End Sub

Private Sub Class_Terminate()
    Set colParam = Nothing
    Set colValue = Nothing
End Sub

Public Function page_Cover() As String
    page_Cover = "cov001.pdg"
End Function

Public Function page_Catalog() As String
    page_Catalog = "!00001.pdg"
End Function

Public Function page_Text(ByVal pageNum As Integer) As String
    page_Text = StrNum(pageNum, 6) & ".pdg"
End Function

Public Sub setHeader(ByVal ht As headerType)
    If ht = toDownload Then
        m_Head = Replace$(m_Head, "e0", "d0")
    Else
        m_Head = Replace$(m_Head, "d0", "e0")
    End If
End Sub

Public Function toString() As String
    Dim sUrl As String

    Dim i As Long
    Dim iEnd As Long
    
    sUrl = m_Head & "?"
    
    'iStart = 1
    iEnd = colParam.Count
    
    For i = 1 To iEnd
        sUrl = sUrl & colParam(i) & "=" & colValue(i) & "&"
    Next
    
    sUrl = Left$(sUrl, Len(sUrl) - 1)
    
    toString = sUrl
End Function

