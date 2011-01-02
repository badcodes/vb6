VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' Shell Extensions - Context Menu Handler
'
' Handler object
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Use at your own risk.
'
' Tested with:
'              * Windows Me/ Windows XP
'              * VB6 SP5
'
' History:
'           08/21/1999 - This code was released
'
'*********************************************************************************************
Option Explicit

' Context menu handlers has to
' implement IContextMenu and IShellExtInit
Implements olelib.IContextMenu
Implements olelib.IShellExtInit

' UDT for store menu item info
Private Type MenuItem
   Caption As String
   HelpStr As String
   Verb As String
   idCMD As Long
   Enable As Boolean
End Type


Const CMDSHELL = 0
Const MSYSSHELL = 1
Const PACKZHTM = 3
Const ZHTMEXSEP = 4
Const ZHTMEXHERE = 5
Const ZHTMOPEN = 2


Private WINRAR As String
Const mnuUbound = 5
Private m_Items(5) As MenuItem
Private m_firstID As Long
'Private mnuUbound As Long
Private m_OldQuery As Long
Private m_SelectedFiles() As String

'Private m_SFolders As New linvblib.CStringCollection
'Private m_SFiles As New linvblib.CStringCollection
'
' IContextMenu::QueryContextMenu
'
' In this function you must add your items
' to the context menu


Friend Function QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As Long) As Long

    'Dim hSubMenu As Long, lIdx As Long
    ' Create a popup menu
    'hSubMenu = CreatePopupMenu()
    On Error GoTo Herr
    Dim idx As Long
    Dim idCMD As Long
    'MSGBOX "hMenu=" & CStr(hMenu)
'    MsgBox "indexMenu=" & CStr(indexMenu)
'    MsgBox "idCmdFirst=" & CStr(idCmdFirst)
'    MsgBox "idCmdLast=" & CStr(idCmdLast)
    'MSGBOX "uFlags=" & CStr(uFlags)
    
    idCMD = idCmdFirst


    InsertMenu hMenu, indexMenu, MF_BYPOSITION Or MF_SEPARATOR, idCMD, ByVal 0&
    idCMD = idCMD + 1
    indexMenu = indexMenu + 1
    
    For idx = 0 To mnuUbound
        'MsgBox m_Items(idx).Caption & ":" & m_Items(idx).IdCmd
        If m_Items(idx).Enable Then
            m_Items(idx).idCMD = idCMD - idCmdFirst
            InsertMenu hMenu, indexMenu, MF_BYPOSITION Or MF_STRING, idCMD, ByVal m_Items(idx).Caption
            indexMenu = indexMenu + 1
            idCMD = idCMD + 1
        End If
    Next
    '    ''MSGBOX  "Out next"
    If idCMD - idCmdFirst > 1 Then
        InsertMenu hMenu, indexMenu, MF_BYPOSITION Or MF_SEPARATOR, idCMD, ByVal 0&
        indexMenu = indexMenu + 1
        idCMD = idCMD + 1
    End If
    ' Return how many non separator
    ' menu items we were added
    QueryContextMenu = idCMD - idCmdFirst
    Exit Function
    
Herr:
    QueryContextMenu = 0
    'MSGBOX "QueryContextMenu Err:" & Err.Description

End Function

'
' Retrieves the language-independent command
' string or the help text for a context
' menu item
'
Private Function getMIndex(ByVal idCMD As Long) As Long
    Dim i As Long
    For i = 0 To mnuUbound
        If m_Items(i).idCMD = idCMD Then
            getMIndex = i
            Exit Function
        End If
    Next
End Function

Private Function FirstID() As Long
FirstID = m_firstID
m_firstID = m_firstID + 1
End Function

Private Sub IContextMenu_GetCommandString(ByVal idCMD As Long, ByVal uType As olelib.GetCommandStringFlags, pwReserved As Long, ByVal pszName As Long, ByVal cchMax As Long)

    On Error GoTo Herr

   idCMD = getMIndex(idCMD)
    'MSGBOX "GetCommandString " & uType & vbCrLf & _
                        "idCmd=" & idCmd & vbCrLf & _
                        "HelpStr=" & m_Items(idCmd).HelpStr
    Select Case uType
        Case GCS_HELPTEXT
            lstrcpynA pszName, m_Items(idCMD).HelpStr, cchMax
        Case (GCS_HELPTEXT Or GCS_UNICODE)
            Dim Unicode() As Byte
            Unicode = m_Items(idCMD).HelpStr & vbNullChar
            MoveMemory ByVal pszName, Unicode(0), IIf(cchMax > UBound(Unicode) + 1, UBound(Unicode) + 1, cchMax)
        Case GCS_VERB
            lstrcpynA pszName, m_Items(idCMD).Verb, cchMax
        Case Else
            lstrcpynA pszName, vbNullChar, cchMax
    End Select
    Exit Sub
Herr:
    'MSGBOX "GETCommandString Err"

End Sub

'
' Carries out the command associated with
' a context menu item.
'
Private Sub IContextMenu_InvokeCommand(lpici As olelib.CMINVOKECOMMANDINFO)

    Dim idx As Long, Verb As String ', Total As Currency
    On Error Resume Next
    
  
    'MsgBox "InvokeCommand"
    'MsgBox lpici.lpVerb
    If mnuUbound < 0 Then Exit Sub
    
    ' Check if lpVerb is an string pointer
    ' or the ID

    If (lpici.lpVerb \ &H10000) <> 0 Then
        ' lpVerb is a string so
        ' copy it from the pointer
        Verb = StrFromPtrA(lpici.lpVerb)
        ' Search the items array for
        ' the command and convert
        ' it to the item index
        For idx = 0 To UBound(m_Items)
            If m_Items(idx).Verb = Verb Then
                Exit For
            End If
        Next
    Else
        idx = lpici.lpVerb
        idx = getMIndex(idx)
    End If
    
    'MSGBOX "Verb=" & Verb
    ' Do the action asociated
    ' with the menu item
    Dim lFUbound As Long
    Dim zhtmName As String
    Dim cmdLine As String
    'Dim FileList As String
    lFUbound = UBound(m_SelectedFiles)
    If lFUbound < 0 Then Exit Sub
    Dim curPath As String
    Dim extrPath As String
    curPath = GetParentFolderName(m_SelectedFiles(0))
    ChDrive Left$(curPath, 1)
    ChDir curPath
    Select Case m_Items(idx).Verb
        Case "pack"
            If TestWinrar = False Then Exit Sub
            '            Dim lzip As cZip
            If Len(curPath) > 3 And lFUbound > 0 Then
                zhtmName = GetBaseName(curPath) & ".zhtm"
            Else
                zhtmName = GetBaseName(m_SelectedFiles(0)) & ".zhtm"
            End If
            '            With lzip
            '            .ZipFile = lzip
            '            .BasePath = curPath
            '            .IncludeSystemAndHiddenFiles = True
            '            .RecurseSubDirs = True
            '            .StoreDirectories = True
            '            .StoreFolderNames = True
            cmdLine = WINRAR & " A -afzip -m5 -r -ep1 "
            cmdLine = cmdLine & Chr$(34) & zhtmName & Chr$(34)
            If lFUbound = 0 Then
               cmdLine = cmdLine & " " & Chr$(34) & GetFileName(m_SelectedFiles(0)) & "\*" & Chr(34)
            Else
            For idx = 0 To lFUbound
                cmdLine = cmdLine & " " & Chr$(34) & GetFileName(m_SelectedFiles(idx)) & Chr(34)
            Next
            End If
            cmdLine = InputBox("将运行以下命令:", "运行", cmdLine)
            If cmdLine <> "" Then Shell cmdLine, vbNormalFocus
        Case "open"
            If TestWinrar = False Then Exit Sub
            zhtmName = m_SelectedFiles(0)
            Shell WINRAR & " " & Chr$(34) & zhtmName & Chr$(34), vbNormalFocus
        Case "extSep"
            If TestWinrar = False Then Exit Sub
            extrPath = BuildPath(GetParentFolderName(m_SelectedFiles(0)), "")
            For idx = 0 To lFUbound
                cmdLine = WINRAR & " X -ad "
                cmdLine = cmdLine & Chr$(34) & m_SelectedFiles(idx) & Chr$(34) & " "
                cmdLine = cmdLine & Chr$(34) & extrPath & Chr$(34)
                ShellAndClose cmdLine, vbNormalFocus
            Next
            '       Case "extIntelligent"
            '            extrPath = BuildPath(GetParentFolderName(m_SelectedFiles(0)), "")
            '            For idx = 0 To lFUbound
            '            cmdLine = WINRAR & " X -ad "
            '            cmdLine = cmdLine & Chr$(34) & m_SelectedFiles(idx) & Chr$(34) & " "
            '            cmdLine = cmdLine & Chr$(34) & extrPath & Chr$(34)
            '            ShellAndClose cmdLine, vbNormalFocus
            '            Next
            '            For idx = 0 To lFUbound
            '            linvblib.LetYourSonPlay extrPath & GetBaseName(m_SelectedFiles(idx))
            '            Next
        Case "extHere"
            If TestWinrar = False Then Exit Sub
            extrPath = BuildPath(GetParentFolderName(m_SelectedFiles(0)), "")
            For idx = 0 To lFUbound
                cmdLine = WINRAR & " X "
                cmdLine = cmdLine & Chr$(34) & m_SelectedFiles(idx) & Chr$(34) & " "
                cmdLine = cmdLine & Chr$(34) & extrPath & Chr$(34)
                ShellAndClose cmdLine, vbNormalFocus
                '            linvblib.UPLonelyFolder Buildpath(GetParentFolderName(m_SelectedFiles(idx)))
            Next
        Case "cmd"
            Dim toWhere As String
            If lFUbound > 0 Then
                toWhere = GetParentFolderName(m_SelectedFiles(0))
            ElseIf (GetAttr(m_SelectedFiles(0)) And vbDirectory) Then
                toWhere = m_SelectedFiles(0)
            Else
                toWhere = GetParentFolderName(m_SelectedFiles(0))
            End If
            If toWhere = "" Then toWhere = Left$(m_SelectedFiles(0), 3)
            Shell "cmd.exe /K CD /D" & toWhere, vbNormalFocus
        Case "msys"
            If lFUbound > 0 Then
                toWhere = GetParentFolderName(m_SelectedFiles(0))
            ElseIf (GetAttr(m_SelectedFiles(0)) And vbDirectory) Then
                toWhere = m_SelectedFiles(0)
            Else
                toWhere = GetParentFolderName(m_SelectedFiles(0))
            End If
            If toWhere = "" Then toWhere = "." 'Left$(m_SelectedFiles(0), 3)
            Shell "loadmsys.bat " & Chr$(34) & toWhere & Chr$(34), vbHide
    End Select

End Sub

Private Sub IContextMenu_QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As olelib.QueryContextMenuFlags)

    ' QueryContextMenu cannot be implemented
    ' this way. So we use ReplaceVTableEntry
    ' to use QueryContextMenu in mdlFunctions.

End Sub

'
' This procedure is called before any other extension interface to pass the selected
' files to the shell extension
'
Private Sub IShellExtInit_Initialize(ByVal pidlFolder As Long, ByVal lpIDataObject As olelib.IDataObject, ByVal hkeyProgID As Long)

    Dim fCount As Long
'    Dim lUbound As Long, lLbound As Long
    Dim l As Long
    Dim extName As String
    Dim allzhtm As Boolean
    
    GetSelectedFiles m_SelectedFiles, fCount, lpIDataObject
    If fCount < 1 Then Exit Sub
    'MsgBox "Fcount=" & fCount
    
'    lLbound = LBound(m_SelectedFiles())
'    lUbound = UBound(m_SelectedFiles())
'    If lUbound - lLbound < 0 Then Exit Sub
    
    allzhtm = True
    For l = 0 To fCount - 1
        extName = LCase$(GetExtensionName(m_SelectedFiles(l)))
        If extName <> "zhtm" Then allzhtm = False: Exit For
        'If extName <> "zhtm" And extName <> "rar" And extName <> "zip" Then allzhtm = False: Exit For
    Next
    
    If allzhtm Then
        m_Items(ZHTMEXHERE).Enable = True ' .idCMD = FirstID()
        m_Items(ZHTMEXSEP).Enable = True '.idCMD = FirstID()
        If fCount = 1 Then
            m_Items(ZHTMEXSEP).Caption = "Extract &To " & GetBaseName(m_SelectedFiles(0))
            m_Items(ZHTMEXSEP).HelpStr = "Extract To " & GetBaseName(m_SelectedFiles(0))
            m_Items(ZHTMOPEN).Enable = True ' .idCMD = FirstID()
        End If
    Else
        m_Items(PACKZHTM).Enable = True ' .idCMD = FirstID()
    End If

End Sub

Private Sub Class_Initialize()

    Dim ICM As IContextMenu
    ' Replace QueryContextMenu entry
    Set ICM = Me
    m_OldQuery = ReplaceVTableEntry(ObjPtr(ICM), 4, AddressOf mdlFunctions.QueryContextMenu)
    Set ICM = Nothing

    ' Initialize menu items array
    
    With m_Items(CMDSHELL)
        .Caption = "Open CMD Shell Here"
        .HelpStr = .Caption
        .Verb = "cmd"
        .idCMD = -1
        .Enable = True
    End With
    With m_Items(MSYSSHELL)
        .Caption = "Open MSYS Shell Here"
        .HelpStr = .Caption
        .Verb = "msys"
        .idCMD = -1
        .Enable = True
    End With
    With m_Items(ZHTMEXHERE)
        .Caption = "&Extract Here"
        .HelpStr = "Extract Here"
        .Verb = "extHere"
        .idCMD = -1
        .Enable = False
    End With
    With m_Items(ZHTMEXSEP)
        .Verb = "extSep"
        .Caption = "Extract &Separately"
        .HelpStr = "Extract Separately"
        .idCMD = -1
        .Enable = False
    End With
    With m_Items(ZHTMOPEN)
        .Caption = "Open by Winrar"
        .HelpStr = .Caption
        .Verb = "open"
        .idCMD = -1
        .Enable = False
    End With
    With m_Items(PACKZHTM)
        .Caption = "&Packed As zhtm"
        .HelpStr = "Packed As zhtm"
        .Verb = "pack"
        .idCMD = -1
        .Enable = False
    End With


End Sub

Private Function TestWinrar() As Boolean
    Dim appIni As String
    appIni = BuildPath(App.Path, "config.ini")
    WINRAR = iniGetSetting(appIni, "Winrar", "Path")
    TestWinrar = True
    If PathExists(WINRAR) = False Then
        WINRAR = InputBox("Need Winrar!" & vbCrLf & "Type the Path Below", "Warning", WINRAR)
        If PathExists(WINRAR) = False Then
            TestWinrar = False
        Else
            iniSaveSetting appIni, "Winrar", "Path", WINRAR
        End If
    End If
End Function

