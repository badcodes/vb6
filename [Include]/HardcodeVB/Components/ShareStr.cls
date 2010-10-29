VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSharedString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorSharedString
    eeBaseSharedString = 13200  ' CSharedString
End Enum

Private h As Long, p As Long

Sub Create(sName As String)
    Dim e As Long
    If sName = sEmpty Then ApiRaise ERROR_BAD_ARGUMENTS
    ' Try to create file mapping of 65535 (only used pages matter)
    h = CreateFileMapping(-1, ByVal pNull, PAGE_READWRITE, _
                          0, 65535, sName)
    ' Save "error" value which may not be an error value
    e = Err.LastDllError
    If h = hNull Then ApiRaise e
    
    ' Get pointer to mapping
    p = MapViewOfFile(h, FILE_MAP_WRITE, 0, 0, 0)
    If p = pNull Then
        CloseHandle h   ' Undo what we did
        ApiRaise Err.LastDllError
    End If
    ' Check cached value to see if new value
    If e <> ERROR_ALREADY_EXISTS Then
        ' Set size of new file mapping by copying 0 to first 4 bytes
        CopyMemory ByVal p, 0, 4
    ' Else
        ' Existing file mapping already initialized
    End If
End Sub

Private Sub Class_Terminate()
    UnmapViewOfFile p
    CloseHandle h
End Sub

' Default property
Property Get item() As String
Attribute item.VB_UserMemId = 0
    If h = hNull Then ErrRaise ERROR_INVALID_DATA
    BugAssert p <> pNull
    ' Copy length out of first 4 bytes of data
    Dim c As Long
    CopyMemory c, ByVal p, 4
    If c Then
        ' Copy the data
        item = String$(c, 0)
        CopyMemoryToStr item, ByVal (p + 4), c * 2
    End If
End Property

Property Let item(s As String)
    If h = hNull Then ErrRaise ERROR_INVALID_DATA
    BugAssert p <> pNull
    Dim c As Long
    c = Len(s)
    ' Copy length to first 4 bytes and string to remainder
    CopyMemory ByVal p, c, 4
    CopyMemoryStr ByVal (p + 4), s, c * 2
End Property
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".SharedString"
        Select Case e
        Case eeBaseSharedString
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

