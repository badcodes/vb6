VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVectorBool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EErrorVectorBool
    eeBaseVectorBool = 13280    ' CVectorBool
End Enum

Private ai() As Long
Private iLast As Long
Private cChunk As Long

Private Sub Class_Initialize()
    cChunk = 1     ' Default size (32 bits) can be overridden
    ReDim Preserve ai(1 To cChunk) As Long
    iLast = 1
End Sub

' Friend properties to make data structure accessible to walker
Friend Property Get Vector(ByVal i As Long) As Boolean
    BugAssert i > 0 And i <= iLast
    ' Compute the array index for bit i
    Dim iIndex As Long
    iIndex = ((i - 1) \ 32) + 1
    Vector = ai(iIndex) And MBytes.Power2(i Mod 32)
End Property

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
' Create a new data walker object and connect to it
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    ' Create a new iterator object
    Dim vectorwalker As CVectorBoolWalker
    Set vectorwalker = New CVectorBoolWalker
    ' Connect it with collection data
    vectorwalker.Attach Me
    ' Return it
    Set NewEnum = vectorwalker.NewEnum
End Function

' Item is the default property
Property Get item(ByVal i As Long) As Boolean
Attribute item.VB_UserMemId = 0
    BugAssert i > 0
    ' Index might fall within the array bounds and still
    ' be greater than iLast. If so, raise an error
    If i > iLast Then ErrRaise eeOutOfBounds
    ' Compute the array index for bit i
    Dim iIndex As Long
    iIndex = ((i - 1) \ 32) + 1
    item = ai(iIndex) And MBytes.Power2(i Mod 32)
End Property

Property Let item(ByVal i As Long, ByVal fItemA As Boolean)
    BugAssert i > 0
    On Error GoTo FailLetItem
    ' Compute the array index for bit i
    Dim iIndex As Long
    iIndex = ((i - 1) \ 32) + 1
    
    If fItemA Then
        ' Set bit i to True
        ai(iIndex) = ai(iIndex) Or MBytes.Power2(i Mod 32)
    Else
        ' Set bit i to False
        ai(iIndex) = ai(iIndex) And Not MBytes.Power2(i Mod 32)
    End If
    If i > iLast Then iLast = i
    Exit Property
FailLetItem:
    If iIndex > UBound(ai) Then
        ' ReDim array to the number of longs needed to
        ' store i bits, plus cChunk longs
        ReDim Preserve ai(1 To iIndex + cChunk) As Long
        Resume      ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Property

Property Get Last() As Long
    Last = iLast
End Property
Property Let Last(iLastA As Long)
    BugAssert iLastA > 0
    ' Compute the array index for bit iLast
    Dim iIndex As Long
    iIndex = ((iLastA - 1) \ 32) + 1
    ' ReDim array to the number of longs needed
    ' to store iLast bits
    ReDim Preserve ai(1 To iIndex) As Long
    iLast = iLastA
End Property

Property Get Chunk() As Long
    ' Return chunk size as number of bits
    Chunk = cChunk * 32
End Property
Property Let Chunk(cChunkA As Long)
    BugAssert cChunkA > 0
    ' Make max chunk size approximately 100 bits (3 * 32)
    Const cMaxChunk = 3
    ' Calculate the number of longs needed to store
    ' cChunkA bits
    Dim cLong As Long
    cLong = ((cChunkA - 1) \ 32) + 1
    ' Store chunk size as a count of longs
    cChunk = IIf(cChunkA < cMaxChunk * 32, cLong, cMaxChunk)
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".VectorBool"
        Select Case e
        Case eeBaseVectorBool
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

