VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenuArrHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public maxItem As Long
Public maxCaptionLength As Long
Private mMenuArray As Object
Public Property Get Menus() As Object
Set Menus = mMenuArray
End Property
Public Property Let Menus(mnuArrObject As Object)
'On Error Resume Next
Set mMenuArray = mnuArrObject
'If mMenuArray.Count < 1 Then mMenuArray(0).Visible = False
End Property

Private Function getShortName(sLongName As String) As String
    Dim lenStr As Long
    Dim lenBase As Long
    Dim lenLeft As Long
    Dim sBase As String
    Dim sLeft As String
    sLongName = toDosPath(sLongName)

    If maxCaptionLength = 0 Then maxCaptionLength = 40
    lenStr = Len(sLongName)

    If lenStr < maxCaptionLength Then
        getShortName = sLongName
        Exit Function
    End If

    sBase = RightRight(sLongName, "\", vbBinaryCompare, ReturnOriginalStr)
    lenBase = Len(sBase)

    If lenBase >= maxCaptionLength Then
        getShortName = sBase
        Exit Function
    End If

    lenLeft = maxCaptionLength - lenBase
    If lenLeft < 6 Then getShortName = sBase: Exit Function
    sLeft = RightLeft(sLongName, "\", vbBinaryCompare, ReturnEmptyStr)
    getShortName = Left$(sLeft, (lenLeft - 2) / 3)
    getShortName = getShortName & "...." & _
       Right$(sLeft, (lenLeft - 2) / 3 * 2)
    getShortName = getShortName & sBase
End Function

Public Sub JustAdd(sCaption As String, Optional sTagKey As String = "")
    
    If maxItem <= 0 Then Exit Sub
    If sTagKey = "" Then sTagKey = sCaption
    
    On Error GoTo errAddMenu
    'Dim iFound As Long
    Dim iIndex As Long
    Dim iCount As Long
    
    iCount = mMenuArray.Count
    If iCount > maxItem Then Exit Sub

    Load mMenuArray(iCount)
    mMenuArray(iCount).Visible = True
    
    For iIndex = iCount To 2 Step -1
        With mMenuArray(iIndex)
        .Caption = mMenuArray(iIndex - 1).Caption
        .Tag = mMenuArray(iIndex - 1).Tag
        End With
    Next
    
    mMenuArray(1).Caption = getShortName(sCaption)
    mMenuArray(1).Tag = sTagKey
    
    Exit Sub
errAddMenu:
    Err.Raise 19811, "CMenuArrHandle::JustAdd", "Unalbe to Add Menu."

End Sub
Public Sub AddUnique(sCaption As String, Optional sTagKey As String = "")
    
    If maxItem <= 0 Then Exit Sub
    If sTagKey = "" Then sTagKey = sCaption
    On Error GoTo errAddMenu
    Dim iFound As Long
    Dim iIndex As Long
    Dim iCount As Long
    iCount = mMenuArray.Count - 1
    
    For iIndex = 1 To iCount
    If StrComp(sTagKey, mMenuArray(iIndex).Tag, vbTextCompare) = 0 Then iFound = iIndex: Exit For
    Next
    
    If iFound = 0 Then
        iCount = iCount + 1
        If iCount > maxItem Then
            iFound = iCount - 1
        Else
            Load mMenuArray(iCount)
            mMenuArray(iCount).Visible = True
            iFound = iCount
        End If
    End If
    
    For iIndex = iFound To 2 Step -1
        With mMenuArray(iIndex)
        .Caption = mMenuArray(iIndex - 1).Caption
        .Tag = mMenuArray(iIndex - 1).Tag
        End With
    Next
    
    mMenuArray(1).Caption = getShortName(sCaption)
    mMenuArray(1).Tag = sTagKey

    Exit Sub
errAddMenu:
    Err.Raise 19811, "CMenuArrHandle::AddUnique", "Unalbe to Add Menu."
End Sub
'Private Function iMenuIndex(sTagKey As String) As Long
'    On Error GoTo 0
'    Dim iIndex As Long, iMcount As Long
'    iMcount = mMenuArray.Count - 1
'    For iIndex = 1 To iMcount
'    If sTagKey = mMenuArray.Tag Then iMenuIndex = iIndex: Exit Function
'    Next
'End Function
Public Sub Clear()
    ReduceTo (1)
End Sub
Public Sub Remove(iIndex As Integer)
On Error GoTo 0
    Dim iLast As Integer
    If iIndex < 1 Then iIndex = 1
    iLast = mMenuArray.Count - 1
    If iIndex > iLast Then Exit Sub
    Do While (iIndex < iLast)
    mMenuArray(iIndex).Caption = mMenuArray(iIndex + 1).Caption
    mMenuArray(iIndex).Tag = mMenuArray(iIndex + 1).Tag
    iIndex = iIndex + 1
    Loop
    Unload mMenuArray(iLast)
End Sub
Public Sub ReduceTo(iIndex As Integer)
    On Error GoTo 0
    Dim iLast As Integer, i As Integer
    If iIndex < 1 Then iIndex = 1
    iLast = mMenuArray.Count - 1
    If iIndex > iLast Then Exit Sub
    For i = iLast To iIndex Step -1
    Unload mMenuArray(i)
    Next
End Sub


Private Sub Class_Terminate()
On Error Resume Next
If mMenuArray.Count > 1 Then
    mMenuArray(0).Visible = True
Else
    mMenuArray(0).Visible = False
End If
'Set mMenuArray = Nothing
End Sub

