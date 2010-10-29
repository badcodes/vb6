VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRecentFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private colRecentFiles As New Collection
Public maxItem As Long
Public maxCaptionLength As Long

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

    sBase = linvblib.GetFileName(sLongName)
    lenBase = Len(sBase)

    If lenBase >= maxCaptionLength Then
        getShortName = sBase
        Exit Function
    End If

    lenLeft = maxCaptionLength - lenBase
    sLeft = linvblib.GetParentFolderName(sLongName)
    getShortName = Left$(sLeft, (lenLeft - 2) / 3)
    getShortName = getShortName & "...." & _
       Right$(sLeft, (lenLeft - 2) / 3 * 2)
    getShortName = getShortName & sBase

End Function

Public Sub Add(sFiletoAdd As String, Optional theKey As String = "")

    If maxItem = 0 Then Exit Sub
    On Error Resume Next

    With colRecentFiles

        If theKey <> "" Then
            .Remove theKey
        Else
            .Remove sFiletoAdd
        End If

        If .Count >= maxItem Then .Remove maxItem

        If .Count > 0 Then
            .Add sFiletoAdd, theKey, 1
        Else
            .Add sFiletoAdd, theKey
        End If

    End With

End Sub

Public Sub FillinMenu(rfMenu As Object)

    Dim l As Long
    Dim iSubMenuCount As Long
    Dim iRecentFilesCount As Long
    On Error Resume Next
    iSubMenuCount = rfMenu.Count - 1
    iRecentFilesCount = colRecentFiles.Count

    If iRecentFilesCount > iSubMenuCount Then

        For l = iSubMenuCount + 1 To iRecentFilesCount
            Load rfMenu(l)
            rfMenu(l).Visible = True
        Next

    Else

        For l = iSubMenuCount To iRecentFilesCount + 1 Step -1
            Unload rfMenu(l)
        Next

    End If

    For l = 1 To iRecentFilesCount
        rfMenu(l).Caption = Replace$(getShortName(colRecentFiles.Item(l)), "|", "\")
        rfMenu(l).Tag = colRecentFiles.Item(l)
    Next

    If iRecentFilesCount > 0 Then
        rfMenu(0).Visible = True
    Else
        rfMenu(0).Visible = False
    End If

End Sub

Public Function Item(Index As Variant) As Variant

    Item = colRecentFiles.Item(Index)

End Function

Public Function Count() As Long

    Count = colRecentFiles.Count

End Function

Public Sub Clear()

    Set colRecentFiles = New Collection

End Sub

'Public Sub SaveToIni(ByVal sIniFilename As String)
'
'    Dim hINi As New CLiNInI
'    hINi.file = sIniFilename
'    Dim l As Long
'    Dim sSecText As String
'
'    With colRecentFiles
'        sSecText = "Count=" & .Count & vbCrLf
'
'        For l = 1 To .Count
'            sSecText = sSecText & Str$(l) & "=" & .Item(l) & vbCrLf
'        Next
'
'    End With
'
'    hINi.SaveSection "RecentFiles", sSecText
'    hINi.Save
'    Set hINi = Nothing
'
'End Sub
'
'Public Sub LoadFromIni(ByVal sIniFilename As String)
'
'    Dim l As Long
'    Dim lCount As Long
'    Dim sf As String
'    Dim hINi As New CLiNInI
'    hINi.file = sIniFilename
'    Set colRecentFiles = New Collection
'    On Error Resume Next
'    lCount = hINi.GetSetting("RecentFiles", "Count")
'
'    If lCount > maxItem Then lCount = maxItem
'
'    For l = 1 To lCount
'        sf = hINi.GetSetting("RecentFiles", Str$(l))
'        colRecentFiles.Add sf, sf
'    Next
'
'    Set hINi = Nothing
'
'End Sub

Public Sub LoadFromMenus(ByRef rfMenu As Object)

    Dim i As Long
    Dim iCount As Long
    Set colRecentFiles = New Collection
    On Error Resume Next
    iCount = rfMenu.Count - 1

    If iCount > maxItem Then iCount = -maxItem

    For i = 1 To iCount

        If rfMenu(i).Tag <> "" Then
            colRecentFiles.Add rfMenu(i).Tag, rfMenu(i).Tag
        End If

    Next

End Sub

