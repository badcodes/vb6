VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVectorInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorVectorInt
    eeBaseVectorInt = 13320     ' CVectorInt
End Enum

Private ai() As Integer
Private iLast As Long
Private cChunk As Long

Private Sub Class_Initialize()
    cChunk = 10     ' Default size can be overridden
    ReDim Preserve ai(1 To cChunk) As Integer
    iLast = 1
End Sub

' Friend properties to make data structure accessible to walker
Friend Property Get Vector(ByVal i As Long) As Integer
    BugAssert i > 0 And i <= iLast
    Vector = ai(i)
End Property

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
' Create a new data walker object and connect to it
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    ' Create a new iterator object
    Dim vectorwalker As CVectorIntWalker
    Set vectorwalker = New CVectorIntWalker
    ' Connect it with collection data
    vectorwalker.Attach Me
    ' Return it
    Set NewEnum = vectorwalker.NewEnum
End Function

' Item is the default property
Property Get item(ByVal i As Long) As Integer
Attribute item.VB_UserMemId = 0
    BugAssert i > 0
    item = ai(i)
End Property

Property Let item(ByVal i As Long, ByVal iItemA As Integer)
    BugAssert i > 0
    On Error GoTo FailLetItem
    ai(i) = iItemA
    If i > iLast Then iLast = i
    Exit Property
FailLetItem:
    If i > UBound(ai) Then
        ReDim Preserve ai(1 To i + cChunk) As Integer
        Resume      ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Property

Property Get Last() As Long
    Last = iLast
End Property
Property Let Last(iLastA As Long)
    BugAssert iLastA > 0
    ReDim Preserve ai(1 To iLastA) As Integer
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
        sSource = App.ExeName & ".VectorInt"
        Select Case e
        Case eeBaseVectorInt
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

