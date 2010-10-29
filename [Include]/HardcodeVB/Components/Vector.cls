VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorVector
    eeBaseVector = 13270    ' CVector
End Enum

Private av() As Variant
Private iLast As Long
Private cChunk As Long

Private Sub Class_Initialize()
    cChunk = 10     ' Default size can be overridden
    ReDim Preserve av(1 To cChunk) As Variant
    iLast = 1
    BugLocalMessage "CVector:Initialize"
End Sub

Private Sub Class_Terminate()
    BugLocalMessage "CVector:Terminate"
End Sub

' Friend properties to make data structure accessible to walker
Friend Property Get Vector(ByVal i As Long) As Variant
    BugAssert i > 0 And i <= iLast
    If IsObject(av(i)) Then
        Set Vector = av(i)
    Else
        Vector = av(i)
    End If
End Property

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
' Create a new data walker object and connect to it
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    ' Create a new iterator object
    Dim vectorwalker As CVectorWalker
    Set vectorwalker = New CVectorWalker
    ' Connect it with collection data
    vectorwalker.Attach Me
    ' Return it
    Set NewEnum = vectorwalker.NewEnum
End Function

' Item is the default property
Property Get item(ByVal i As Long) As Variant
Attribute item.VB_UserMemId = 0
    BugAssert i > 0
    ' If index is out-of-range, return default (Empty)
    On Error Resume Next
    If IsObject(av(i)) Then
        Set item = av(i)
    Else
        item = av(i)
    End If
End Property

Property Let item(ByVal i As Long, ByVal vItemA As Variant)
    BugAssert i > 0
    On Error GoTo FailLetItem
    av(i) = vItemA
    If i > iLast Then iLast = i
    Exit Property
FailLetItem:
    If i > UBound(av) Then
        ReDim Preserve av(1 To i + cChunk) As Variant
        Resume              ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Property

Property Set item(ByVal i As Long, ByVal vItemA As Variant)
    BugAssert i > 0
    On Error GoTo FailSetItem
    Set av(i) = vItemA
    If i > iLast Then iLast = i
    Exit Property
FailSetItem:
    If i > UBound(av) Then
        ReDim Preserve av(1 To i + cChunk) As Variant
        Resume      ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Property

Property Get Last() As Long
    Last = iLast
End Property
Property Let Last(iLastA As Long)
    BugAssert iLastA > 0
    ReDim Preserve av(1 To iLastA) As Variant
    iLast = iLastA
End Property

Property Get Chunk() As Long
    Chunk = cChunk
End Property
Property Let Chunk(cChunkA As Long)
    BugAssert cChunkA > 0
    cChunk = IIf(cChunkA < 100, cChunkA, 100)
End Property
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Vector"
        Select Case e
        Case eeBaseVector
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

