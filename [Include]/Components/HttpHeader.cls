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
Attribute VB_Exposed = True
Option Explicit
Private mRespone As String
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
    mRespone = ""
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
            mRespone = StripString(sHeaderMap(0), " ")
            mStatus = CLng(Left$(mRespone, 3))
            mRespone = Mid$(mRespone, 5)
        End If
        
        iFirst = 1
    Else
        iFirst = 0
    End If
    
    
    
    Dim iPos As Long
    Dim sFieldName As String
    Dim sFieldData As String
    Dim i As Long
    i = 1
    For iStart = iFirst To iEnd
        i = i + 1
        iPos = 0
        iPos = InStr(sHeaderMap(iStart), ": ")
        If iPos > 1 Then
            sFieldName = Left$(sHeaderMap(iStart), iPos - 1)
            sFieldData = Mid$(sHeaderMap(iStart), iPos + 2)
            mHeaderMap(i, 1) = sFieldName
            mHeaderMap(i, 2) = sFieldData
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

Public Property Get HttpRespone() As String
    HttpRespone = mRespone
End Property

Public Property Let HttpRespone(ByVal vText As String)
    mRespone = vText
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

    Dim i As Long
    
    For i = 2 To 100
        If (mHeaderMap(i, 1) <> "") Then
            HeaderString = HeaderString & vbCrLf & mHeaderMap(i, 1) & ": " & mHeaderMap(i, 2)
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
    Dim i As Long
        For i = 2 To 100
            If mHeaderMap(i, 1) = "" And mHeaderMap(i, 2) = "" Then
                SearchFree = i
                Exit Function
            End If
        Next
End Function

Private Function Search(ByVal vName As String) As Long
    Dim i As Long
    For i = 1 To 100
        If (StrComp(vName, mHeaderMap(i, 1), vbTextCompare) = 0) Then
            Search = i
            Exit Function
        End If
    Next
End Function
Public Function GetField(ByVal vName As String) As String
    If vName = "" Then
        GetField = mHeaderMap(1, 1)
    Else
        Dim i As Long
        i = Search(vName)
        If (i > 0) Then GetField = mHeaderMap(i, 2)
    End If
End Function

Public Sub DeleteField(ByVal vName As String)
        If vName = "" Then
            mHeaderMap(1, 1) = ""
            Exit Sub
        End If
        Dim i As Long
        i = Search(vName)
        If (i > 0) Then mHeaderMap(i, 1) = ""
End Sub

Public Sub SetField(ByRef vName As String, ByRef vData As String)
Dim i As Long
If vName = "" Then
    mHeaderMap(1, 1) = vData
    Exit Sub
End If
i = Search(vName)
If i > 0 Then
    mHeaderMap(i, 2) = vData
Else
    i = SearchFree()
    If i > 0 Then
        mHeaderMap(i, 1) = vName
        mHeaderMap(i, 2) = vData
    End If
End If

End Sub


