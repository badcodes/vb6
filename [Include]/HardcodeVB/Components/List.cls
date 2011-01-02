VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorList
    eeBaseList = 13100  ' CList
End Enum

Private lnkHead As CLink
Private c As Long

Private Sub Class_Initialize()
    BugLocalMessage "CList:Initialize"
End Sub

Private Sub Class_Terminate()
    BugLocalMessage "CList:Terminate"
End Sub

' Make data structure available to cooperating classes
Friend Property Get Head() As CLink
    Set Head = lnkHead
End Property

' Insert at head of list
Sub Add(vItem As Variant)
    ' Create temporary link with new value
    Dim lnkTmp As CLink
    Set lnkTmp = New CLink
    If IsObject(vItem) Then
        Set lnkTmp.item = vItem
    Else
        lnkTmp.item = vItem
    End If
    ' Point it where previous head pointed
    Set lnkTmp.NextLink = lnkHead
    ' Attach it to front
    Set lnkHead = lnkTmp
    ' lnkTmp temporary goes out of scope and disappears
    c = c + 1
End Sub

Sub Remove(Optional vIndex As Variant = 1)
    ' Can't remove from empty list
    If lnkHead Is Nothing Then Exit Sub
    ' Walk through to find the item
    Dim i As Long, lnkTmp As CLink, walker As New CListWalker
    Dim iIndex As Long, sIndex As String

    walker.Attach Me
    ' Save last link for unhooking current
    Set lnkTmp = lnkHead
    ' Find the matching link
    If VarType(vIndex) = vbString Then
        ' Remove by string key (ignore if no string compare)
        sIndex = vIndex
        Do While walker.More
            On Error Resume Next
            With walker.CurLink
                If .item = sIndex Then
                    If walker.CurLink Is lnkHead Then
                        ' First can be deleted only by changing head
                        Set lnkHead = .NextLink
                    Else
                        ' Delete matching link by hooking
                        ' its next to previous
                        Set lnkTmp.NextLink = .NextLink
                    End If
                    c = c - 1
                    Exit Sub
                End If
            End With
            ' Save last link for unhooking current
            Set lnkTmp = walker.CurLink
        Loop
    Else
        ' Remove by numeric index
        iIndex = vIndex
        Do While walker.More
            i = i + 1
            If iIndex = i Then
                With walker.CurLink
                    If i = 1 Then
                        ' First can be deleted only by changing head
                        Set lnkHead = .NextLink
                    Else
                        ' Delete matching link by hooking
                        ' its next to previous
                        Set lnkTmp.NextLink = .NextLink
                    End If
                    c = c - 1
                    Exit Sub
                End With
            End If
            ' Save last link for unhooking current
            Set lnkTmp = walker.CurLink
        Loop
    End If
    ' No match found
End Sub

Property Get Count() As Long
    Count = c
End Property

' Remove all items
Sub Clear()
    If lnkHead Is Nothing Then Exit Sub
    Do Until lnkHead.NextLink Is Nothing
        Set lnkHead.NextLink = lnkHead.NextLink.NextLink
    Loop
    Set lnkHead = Nothing
    c = 0
End Sub

' Default property
Property Get item(Optional vIndex As Variant = 1) As Variant
Attribute item.VB_UserMemId = 0
    If lnkHead Is Nothing Then Exit Property
    ' Walk through to find the item
    Dim walker As New CListWalker, v As Variant
    Dim i As Long, iIndex As Long, sIndex As String
    
    ' Find the matching link
    walker.Attach Me
    If VarType(vIndex) = vbString Then
        ' Search by string key
        sIndex = vIndex
        ' Ignore error for entries that can't be string compared
        On Error Resume Next
        Do While walker.More
            With walker.CurLink
                If .item = sIndex Then
                    If IsObject(.item) Then
                        Set item = .item
                    Else
                        item = .item
                    End If
                End If
            End With
        Loop
    Else
        ' Search by numeric index
        iIndex = vIndex
        Do While walker.More
            i = i + 1
            With walker.CurLink
                If iIndex = i Then
                    If IsObject(.item) Then
                        Set item = .item
                    Else
                        item = .item
                    End If
                End If
            End With
        Loop
    End If
    ' Item = Empty
End Property

Property Let item(Optional vIndex As Variant = 1, vItemA As Variant)
    If lnkHead Is Nothing Then Exit Property
    ' Walk through to find the item
    Dim walker As New CListWalker, v As Variant
    Dim i As Long, iIndex As Long, sIndex As String
    ' Check type outside loop
    If VarType(vIndex) = vbString Then
        sIndex = vIndex
        iIndex = -1
    Else
        iIndex = vIndex
    End If
    
    ' Find the matching link
    walker.Attach Me
    Do While walker.More
        i = i + 1
        With walker.CurLink
            If iIndex = -1 Then
                ' Ignore error for entries that can't be string compared
                On Error Resume Next
                If .item = sIndex Then .item = vItemA
                On Error GoTo 0
            Else
                If CLng(vIndex) = i Then .item = vItemA
            End If
        End With
    Loop
    ' Item = Empty
End Property

Property Set item(Optional vIndex As Variant = 1, vItemA As Variant)
    If lnkHead Is Nothing Then Exit Property
    ' Walk through to find the item
    Dim walker As New CListWalker, v As Variant
    Dim i As Long, iIndex As Long, sIndex As String
    ' Check type outside loop
    If VarType(vIndex) = vbString Then
        sIndex = vIndex
        iIndex = -1
    Else
        iIndex = vIndex
    End If
    
    ' Find the matching link
    walker.Attach Me
    Do While walker.More
        i = i + 1
        With walker.CurLink
            If iIndex = -1 Then
                ' Ignore error for entries that can't be string compared
                On Error Resume Next
                If .item = sIndex Then Set .item = vItemA
                On Error GoTo 0
            Else
                If CLng(vIndex) = i Then Set .item = vItemA
            End If
        End With
    Loop
    ' Item = Empty
End Property

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
' Create a new data walker object and connect to it
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    ' Create a new iterator object
    Dim listwalker As CListWalker
    Set listwalker = New CListWalker
    ' Connect it with collection data
    listwalker.Attach Me, True
    ' Return it
    Set NewEnum = listwalker.NewEnum
End Function

'
#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".List"
        Select Case e
        Case eeBaseList
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

