VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHttpHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mResponse As String
Private mStatus As Long
Private mDate As String
Private mModified As String
Private mType As String
Private mSize As Long
Private mHttpGet
Private mLocation As String
Private mHeaderMap(1 To 100, 1 To 2) As String


Public Sub Init(HeaderString As String)
    On Error Resume Next
    mResponse = ""
    mStatus = -1
    mType = ""
    mSize = 0
    'mHeader = ""
    
    If HeaderString = "" Then Exit Sub
    
    Dim iStart As Long
    Dim iEnd As Long
    Dim sHeaderMap() As String
    HeaderString = Replace$(HeaderString, Chr$(0), vbCrLf)
    sHeaderMap = Split(HeaderString, vbCrLf)
    iEnd = ArrayUbound(sHeaderMap)
    
    If iEnd < 0 Then Exit Sub
    Dim iFirst As Long
    
    If InStr(sHeaderMap(0), ": ") < 1 Then
        mHeaderMap(1, 1) = sHeaderMap(0)
        
        
        If Left$(sHeaderMap(0), 3) = "GET" Then
            mHttpGet = StripString(sHeaderMap(0), " ", " ")
        Else
            mResponse = StripString(sHeaderMap(0), " ")
            mStatus = CLng(Left$(mResponse, 3))
            mResponse = Mid$(mResponse, 5)
        End If
        
        iFirst = 1
    Else
        iFirst = 0
    End If
    
    
    
    Dim iPos As Long
    Dim sFieldName As String
    Dim sFieldData As String
    Dim I As Long
    I = 1
    For iStart = iFirst To iEnd
        I = I + 1
        iPos = 0
        iPos = InStr(sHeaderMap(iStart), ": ")
        If iPos > 1 Then
            sFieldName = Left$(sHeaderMap(iStart), iPos - 1)
            sFieldData = Mid$(sHeaderMap(iStart), iPos + 2)
            mHeaderMap(I, 1) = sFieldName
            mHeaderMap(I, 2) = sFieldData
            Select Case UCase$(sFieldName)
            
                Case "DATE"
                    'mDate = CDate(sFieldData)
                    mDate = sFieldData
                Case "CONTENT-TYPE"
                    mType = sFieldData
                Case "CONTENT-LOCATION"
                    mLocation = sFieldData
                Case "CONTENT-LENGTH"
                    mSize = CLng(sFieldData)
                Case "LAST-MODIFIED"
                    'mModified = CDate(sFieldData)
                    mModified = sFieldData
            End Select
        Else
            mHeaderMap(iStart + 1, 1) = sHeaderMap(iStart)
        End If
    Next
    
    
    
End Sub

Private Function LeftString(ByRef vString As String, ByRef vFind As String, Optional Start As Long = 1) As String
    Dim iEnd As Long
    iEnd = InStr(Start, vString, vFind)
    If iEnd < 1 Then Exit Function
    LeftString = Left$(vString, iEnd - 1)
    
End Function

Private Function StripString(ByRef vString As String, ByRef vStart As String, Optional ByRef vEnd As String = vbNullString, Optional ByVal Start As Long = 1) As String
Dim iStart As Long
Dim iEnd As Long

iStart = InStr(Start, vString, vStart)
If iStart < 1 Then Exit Function

iStart = iStart + Len(vStart)
If vEnd = vbNullString Then
    StripString = Mid$(vString, iStart)
    Exit Function
End If

iEnd = InStr(iStart, vString, vEnd)
If iEnd < 1 Then Exit Function

StripString = Mid$(vString, iStart, iEnd - iStart)


End Function

Public Property Get HttpGet() As String
    HttpGet = mHttpGet
End Property

Public Property Get HttpResponse() As String
    HttpResponse = mResponse
End Property

Public Property Let HttpResponse(ByVal vText As String)
    mResponse = vText
End Property

Public Property Get HttpStatus() As Long
    HttpStatus = mStatus
End Property
Public Property Let HttpStatus(ByVal vStatus As Long)
    mStatus = vStatus
End Property
Public Property Get ConentLength() As Long
    ConentLength = mSize
End Property

Public Property Get LastModified() As Date
    LastModified = mModified
End Property

Public Property Get FileDate() As Date
    FileDate = mDate
End Property

Public Property Get ContentType() As String
    ContentType = mType
End Property

Public Property Let ContentType(ByVal vType As String)
    mType = vType
End Property
Public Property Get ContentLocation() As String
    ContentLocation = mLocation
End Property
Public Property Get HeaderString() As String
    If mHeaderMap(1, 2) <> "" And mHeaderMap(1, 1) <> "" Then
        HeaderString = mHeaderMap(1, 1) & ": " & mHeaderMap(1, 2)
    ElseIf mHeaderMap(1, 1) <> "" Then
        HeaderString = mHeaderMap(1, 1)
    End If

    Dim I As Long
    
    For I = 2 To 100
        If (mHeaderMap(I, 1) <> "") Then
            HeaderString = HeaderString & vbCrLf & mHeaderMap(I, 1) & ": " & mHeaderMap(I, 2)
        End If
    Next
    
    If Left$(HeaderString, Len(vbCrLf)) = vbCrLf Then HeaderString = Mid$(HeaderString, 3)
End Property

Private Function ArrayUbound(ByRef vArr() As String) As Long
    On Error Resume Next
    
    ArrayUbound = -2
    ArrayUbound = UBound(vArr())
    
End Function

Private Function ArrayLbound(ByRef vArr() As String) As Long
On Error Resume Next
    ArrayLbound = -1
    ArrayLbound = LBound(vArr())
End Function

Private Function SearchFree() As Long
    Dim I As Long
        For I = 2 To 100
            If mHeaderMap(I, 1) = "" And mHeaderMap(I, 2) = "" Then
                SearchFree = I
                Exit Function
            End If
        Next
End Function

Private Function Search(ByVal vName As String) As Long
    Dim I As Long
    For I = 1 To 100
        If (StrComp(vName, mHeaderMap(I, 1), vbTextCompare) = 0) Then
            Search = I
            Exit Function
        End If
    Next
End Function
Public Function GetField(ByVal vName As String) As String
    If vName = "" Then
        GetField = mHeaderMap(1, 1)
    Else
        Dim I As Long
        I = Search(vName)
        If (I > 0) Then GetField = mHeaderMap(I, 2)
    End If
End Function

Public Sub DeleteField(ByVal vName As String)
        If vName = "" Then
            mHeaderMap(1, 1) = ""
            Exit Sub
        End If
        Dim I As Long
        I = Search(vName)
        If (I > 0) Then mHeaderMap(I, 1) = ""
End Sub

Public Sub SetField(ByRef vName As String, ByRef vData As String)
Dim I As Long
If vName = "" Then
    mHeaderMap(1, 1) = vData
    Exit Sub
End If
I = Search(vName)
If I > 0 Then
    mHeaderMap(I, 2) = vData
Else
    I = SearchFree()
    If I > 0 Then
        mHeaderMap(I, 1) = vName
        mHeaderMap(I, 2) = vData
    End If
End If

End Sub

