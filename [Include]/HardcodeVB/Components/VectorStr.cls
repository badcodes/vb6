VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVectorStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorVectorStr
    eeBaseVectorStr = 13380     ' CVectorStr
End Enum

Private astr() As String
Private iLast As Long
Private cChunk As Long

Private Sub Class_Initialize()
    cChunk = 10     ' Default size can be overridden
    ReDim Preserve astr(1 To cChunk) As String
    iLast = 1
End Sub

' Friend properties to make data structure accessible to walker
Friend Property Get Vector(ByVal i As Long) As String
    BugAssert i > 0 And i <= iLast
    Vector = astr(i)
End Property

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
' Create a new data walker object and connect to it
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    ' Create a new iterator object
    Dim vectorwalker As CVectorStrWalker
    Set vectorwalker = New CVectorStrWalker
    ' Connect it with collection data
    vectorwalker.Attach Me
    ' Return it
    Set NewEnum = vectorwalker.NewEnum
End Function

' Item is the default property
Property Get item(ByVal i As Long) As String
Attribute item.VB_UserMemId = 0
    BugAssert i > 0
    item = astr(i)
End Property

Property Let item(ByVal i As Long, ByVal sItemA As String)
    BugAssert i > 0
    On Error GoTo FailLetItem
    astr(i) = sItemA
    If i > iLast Then iLast = i
    Exit Property
FailLetItem:
    If i > UBound(astr) Then
        ReDim Preserve astr(1 To i + cChunk) As String
        Resume      ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Property

Property Get Last() As Long
    Last = iLast
End Property
Property Let Last(iLastA As Long)
    BugAssert iLastA > 0
    ReDim Preserve astr(1 To iLastA) As String
    iLast = iLastA
End Property

Property Get Chunk() As Long
    Chunk = cChunk
End Property
Property Let Chunk(cChunkA As Long)
    BugAssert cChunkA > 0
    cChunk = IIf(cChunkA < 100, cChunkA, 100)
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".VectorStr"
        Select Case e
        Case eeBaseVectorStr
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

