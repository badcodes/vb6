Attribute VB_Name = "MFoldTool"
Option Explicit

Public Enum EErrorFoldTool
    eeBaseFoldTool = 13500      ' FoldTool
End Enum

Public Enum EWalkMode
    ewmFolders = SHCONTF_FOLDERS
    ewmNonfolders = SHCONTF_NONFOLDERS
    ewmBoth = SHCONTF_FOLDERS Or SHCONTF_NONFOLDERS
    ewmIncludeHidden = SHCONTF_INCLUDEHIDDEN
End Enum

Private iidShellFolder As UUID
Private iidContextMenu As UUID
#If fComponent = 0 Then
Private fInitialized As Boolean
#End If

Sub Class_Initialize()
    ' Initialize GUID constants from strings
    IIDFromString "{000214E6-0000-0000-C000-000000000046}", iidShellFolder
    IIDFromString "{000214E4-0000-0000-C000-000000000046}", iidContextMenu
#If fComponent = 0 Then
    ' Set initialized flag for standard module
    fInitialized = True
#End If
End Sub

Private Sub InitIf()
#If fComponent = 0 Then
    If Not fInitialized Then Class_Initialize
#End If
End Sub

'' Item ID helpers (error handling improved in VB6 update)

' Get the byte size of an ID list
Function ItemIDSize(ByVal pidl As Long) As Integer
    If pidl Then CopyMemory ItemIDSize, ByVal pidl, 2
End Function

' Counts the item IDs in an ID list
Function PidlCount(ByVal pidl As Long) As Long
    Dim cItem As Long
    If pidl = pNull Then Exit Function
    Do While ItemIDSize(pidl)
        pidl = NextItemID(pidl)
        cItem = cItem + 1
    Loop
    PidlCount = cItem
End Function

' Get the next item ID in an item ID list
Function NextItemID(ByVal pidl As Long) As Long
    Dim c As Integer
    If pidl = pNull Then Exit Function
    c = ItemIDSize(pidl)
    If c = 0 Then Exit Function
    NextItemID = pidl + c
End Function

' Duplicate an item ID (creator must free with allocator)
Function DuplicateItemID(pidl As Long) As Long
    Dim c As Integer, pidlNew As Long, iZero As Integer ' = 0
    If pidl = pNull Then Exit Function
    ' Get the size
    c = ItemIDSize(pidl)
    If c = 0 Then Exit Function
    ' Allocate space plus two for zero terminator
    On Error Resume Next
    pidlNew = Allocator.alloc(c + 2)
    If pidlNew = pNull Then Exit Function
    
    ' Copy the pidl data
    CopyMemory ByVal pidlNew, ByVal pidl, c
    ' Terminating zero
    CopyMemory ByVal pidlNew + c, iZero, 2
    DuplicateItemID = pidlNew
End Function

' Concatenate two item IDs
Function DuplicateItemIDs(ByVal pidl1 As Long, _
                          ByVal pidl2 As Long) As Long
    Dim pidlNew As Long, cb1 As Long, cb2 As Long, iZero As Integer ' = 0
    ' May be NULL
    If pidl1 Then
        cb1 = ItemIDSize(pidl1)
        If cb1 Then cb1 = cb1 - 2
    End If
    cb2 = ItemIDSize(pidl2)

    ' Allocate space plus two for zero terminator
    pidlNew = Allocator.alloc(cb1 + cb2 + 2)
    If pidlNew Then
        If pidl1 Then
            CopyMemory ByVal pidlNew, ByVal pidl1, cb1
            If pidl2 Then
                CopyMemory ByVal pidlNew + cb1, ByVal pidl2, cb2
                CopyMemory ByVal pidlNew + cb1 + cb2, iZero, 2
            Else
                CopyMemory ByVal pidlNew + cb1, iZero, 2
            End If
        End If
    End If
    DuplicateItemIDs = pidlNew
End Function

Function GetDesktopFolder() As IVBShellFolder
    SHGetDesktopFolder GetDesktopFolder
End Function

Function PathFromPidl(ByVal pidl As Long) As String
    Dim s As String, f As Long
    s = String$(cMaxPath, 0)
    f = SHGetPathFromIDList(pidl, s)
    If f Then PathFromPidl = MUtility.StrZToStr(s)
End Function

Function NameFromPidl(ByVal pidl As Long) As String
    Dim shfi As SHFILEINFO, f As Long
    f = SHGetItemInfo(pidl, 0, shfi, LenB(shfi), _
                      SHGFI_DISPLAYNAME Or SHGFI_PIDL)
    If f Then NameFromPidl = MBytes.ByteZToStr(shfi.szDisplayName)
End Function

Function PidlFromPath(sPath As String) As Long
    Dim pidl As Long, f As Long
    f = SHGetPathFromIDList(pidl, sPath)
    If f Then PidlFromPath = pidl
End Function

Function PathToPidl(sPath As String) As Long
    InitIf
    
    Dim folder As IVBShellFolder
    Dim pidlMain As Long, cParsed As Long, afItem As Long

    ' Make sure the file name is fully qualified
    sPath = MUtility.GetFullPath(sPath)

    ' Convert the path name into a pointer to an item ID list (pidl)
    Set folder = GetDesktopFolder
    folder.ParseDisplayName hNull, 0, sPath, cParsed, pidlMain, afItem
    
    PathToPidl = pidlMain
    
End Function

Function ToPidl(ByVal i As Long) As Long
    ' Set of imaginable special folder constant
    If i >= CSIDL_DESKTOP And (i <= 32767) Then
        ToPidl = PidlFromSpecialFolder(i)
    Else
        ToPidl = i
    End If
End Function

Function PidlFromSpecialFolder( _
                Optional ByVal csidl As ECSIDL = CSIDL_DESKTOP, _
                Optional ByVal hWnd As Long = hNull) As Long
    InitIf  ' Initialize if in standard modue
    On Error Resume Next
    Dim pidl As Long
    
    SHGetSpecialFolderLocation hWnd, csidl, pidl
    If Err = 0 Then PidlFromSpecialFolder = pidl
End Function

' Get folder and pidl from an item (path, pidl, or special folder)
' Note that caller owns any returned pidl and should free it
Function FolderFromItem(vItem As Variant, _
                        Optional pidl As Long = -1) As IVBShellFolder
    InitIf  ' Initialize if in standard modue
    
    Dim folder As IVBShellFolder, folderNext As IVBShellFolder
    Dim pidlItem As Long, pidlTmp As Long, cItem As Long
   
    On Error GoTo FolderFromItemFail
    Set folder = GetDesktopFolder
    If VarType(vItem) = vbString Then
        ' Make sure the file name is fully qualified
        vItem = MUtility.GetFullPath(CStr(vItem))
    
        ' Convert path name to pointer to an item ID list (pidl)
        Dim cParsed As Long, afItem As Long
        folder.ParseDisplayName hNull, 0, CStr(vItem), _
                                cParsed, pidlItem, afItem
    Else
        ' If necessary, convert special folder to pidl
        pidlItem = ToPidl(vItem)
    End If

    ' Walk the list of item IDs and bind to each subfolder in list
    ' to find the folder containing the specified pidl

    cItem = PidlCount(pidlItem)
    ' If caller requests a pidl return, adjust to return pidl of parent
    If pidl <> -1 Then cItem = cItem - 1
    Do While cItem

        ' Create a one-item ID list for the next item in pidlMain
        pidlTmp = DuplicateItemID(pidlItem)
        If pidlTmp = pNull Then GoTo FolderFromItemFail

        Debug.Print GetFolderName(folder, pidlTmp, SHGDN_NORMAL)
        
        ' Bind to the folder specified in the new igtem ID list
        folder.BindToObject pidlTmp, 0, _
                            iidShellFolder, folderNext
        cItem = cItem - 1
        
        ' Release parent folder and reference current child
        Set folder = folderNext
        ' Free temporary pidl
        Allocator.Free pidlTmp
        pidlTmp = pNull
        ' Point to next item (if any)
        If cItem Then pidlItem = NextItemID(pidlItem)
    Loop
    Set FolderFromItem = folder
    If pidl = -1 Then
        ' Free temporary pidl if user doen't request it
        Allocator.Free pidlItem
    Else
        ' User who asked for pidl must free it
        pidl = pidlItem
    End If
    pidl = pidlItem
    Exit Function
    
FolderFromItemFail:
    pidl = pNull
    If pidlTmp <> pNull Then Allocator.Free pidlTmp
                   
End Function

' Use structure from hell to get a folder name from one of three formats
Function GetFolderName(folder As IVBShellFolder, ByVal pidl As Long, _
                       ByVal gdn As ESHGDN) As String
    InitIf  ' Initialize if in standard modue

    Dim s As String, p As Long, c As Long
    Dim ab() As Byte, typefromhell As STRRET
    On Error Resume Next
    folder.GetDisplayNameOf pidl, gdn, typefromhell
    If Err Then Err.Raise Err, "VBCore.FoldTool", ApiError(Err)
    Select Case typefromhell.uType
    Case STRRET_WSTR
        ' Pointer to Unicode string (in first four bytes of byte array)
        CopyMemory p, ByVal VarPtr(typefromhell.CStr(0)), 4
        ' Use Unicode-specific conversion function
        s = MUtility.UPointerToString(p)
        Allocator.Free p
        
    Case STRRET_OFFSET
        ' Offset (in four bytes of byte array) from pidl to ANSI string
        CopyMemory p, ByVal VarPtr(typefromhell.CStr(0)), 4
        p = pidl + p
        ' Use ANSI-specific conversion function
        s = MUtility.APointerToString(p)
        
    Case STRRET_CSTR
        ' ANSI string in byte array
        s = MBytes.AByteZToStr(typefromhell.CStr)

    End Select
    GetFolderName = s
End Function

Function BindToShell(folder As IVBShellFolder, _
                     ByVal pidl As Long) As IVBShellFolder
    Dim folderNew As IVBShellFolder
    folder.BindToObject pidl, 0, iidShellFolder, folderNew
    Set BindToShell = folderNew
End Function

Function FileInfoFromFolder(folder As IVBShellFolder, _
                            ByVal pidl As Long) As CFileInfo
    Dim gao As ESFGAO, sName As String
    Static fi As New CFileInfo
    gao = SFGAO_FILESYSTEM
    ' Determine what type of object you have
    folder.GetAttributesOf 1, pidl, gao
    If gao And SFGAO_FILESYSTEM Then
        ' Use folder parsing name to get file data
        Dim fd As WIN32_FIND_DATA, h As Long
        ' GetFolderName does horrible stuff with STRRET
        sName = GetFolderName(folder, pidl, SHGDN_FORPARSING)
        ' Handle drives
        If Len(sName) = 3 Then
            If Mid$(sName, 2, 2) = ":\" Then
                Dim drive As CDrive
                Set drive = New CDrive
                drive = sName
                With drive
                    fi.CreateFromDrive .Root, .KindStr, _
                                       CCur(.FreeBytes), CCur(.TotalBytes)
                End With
                Set FileInfoFromFolder = fi
                Exit Function
            End If
        End If
        ' Handle files
        h = FindFirstFile(sName, fd)
        If h <> hInvalid Then
            FindClose h
            fi.CreateFromFile sName, fd.dwFileAttributes, _
                              fd.nFileSizeLow, fd.ftLastWriteTime, _
                              fd.ftLastAccessTime, fd.ftCreationTime
            Set FileInfoFromFolder = fi
            Exit Function
        End If
    Else
        ' Some folders don't work with SHGetFileInfo, but GetFolderName works
        sName = GetFolderName(folder, pidl, SHGDN_NORMAL)
    End If
    fi.CreateFromNamePidl sName, pidl
    Set FileInfoFromFolder = fi
End Function

Function WalkAllFolders(folder As IVBShellFolder, foldit As IUseFolder, _
                        Optional ByVal Level As Long = 0, _
                        Optional ByVal ewm As EWalkMode = ewmBoth, _
                        Optional ByVal hWnd As Long = hNull) As Long
    InitIf  ' Initialize if in standard module
    
    Dim idenum As IVBEnumIDList, folderNew As IVBShellFolder
    Dim pidl As Long, cFetched As Long, afAttrib As Long
    
    ' Get the IEnumIDList object for the given folder
    On Error GoTo WalkAllFoldersFail
    folder.EnumObjects hWnd, ewm, idenum
    
    ' Enumerate through the list of folder and nonfolder objects
    On Error GoTo WalkAllFoldersFail2
    Dim hRes As Long
    Do
        hRes = idenum.Next(1, pidl, cFetched)
        ' 0 means got another, 1 means no more, anything else is error
        ' but there had better not be any errors because we'll ignore them
        If hRes Then Exit Do
        
        ' Pass to user-implemented interface to do something with folder
        ' (True in return means user requested termination)
        WalkAllFolders = foldit.UseFolder(Level, folder, pidl)
        If WalkAllFolders Then
            Allocator.Free pidl
            Exit Function
        End If
        
        ' It's not in the docs, but you pass in the attributes you want
        ' to check and GetAttributes passes back whether those attributes
        ' are set, ignoring all others
        afAttrib = SFGAO_HASSUBFOLDER Or SFGAO_FOLDER
        folder.GetAttributesOf 1, pidl, afAttrib
        
        ' If there are subfolders, process them recursively
        If afAttrib And (SFGAO_HASSUBFOLDER Or SFGAO_FOLDER) Then
            folder.BindToObject pidl, 0, iidShellFolder, folderNew
            WalkAllFolders = WalkAllFolders(folderNew, foldit, Level + 1, ewm)
        End If
WalkAllFoldersFail2:
        ' Free the pidl from Next
        Allocator.Free pidl
    Loop
WalkAllFoldersFail:

End Function

Function WalkFolders(folder As IVBShellFolder, foldit As IUseFolder, _
                     Optional UserData As Variant, _
                     Optional ByVal ewm As EWalkMode = ewmBoth, _
                     Optional ByVal hWnd As Long = hNull) As Long
    InitIf  ' Initialize if in standard modue
    
    Dim idenum As IVBEnumIDList, folderNew As IVBShellFolder
    Dim pidl As Long, cFetched As Long, afAttrib As Long
    
    ' Get the IEnumIDList object for the given folder
    On Error GoTo WalkFoldersFail
    folder.EnumObjects hWnd, ewm, idenum
    
    ' Enumerate through the list of folder and nonfolder objects
    On Error GoTo WalkFoldersFail2
    Dim hRes As Long
    Do
        hRes = idenum.Next(1, pidl, cFetched)
        ' 0 means got another, 1 means no more, anything else is error
        ' but there had better not be any errors because we'll ignore them
        If hRes Then Exit Do
        
        ' Pass to user-implemented interface to do something with folder
        ' (True in return means user requested termination)
        WalkFolders = foldit.UseFolder(UserData, folder, pidl)
        If WalkFolders Then
            Allocator.Free pidl
            Exit Function
        End If
        
WalkFoldersFail2:
        ' Free the pidl from Next
        Allocator.Free pidl
    Loop
    Exit Function
WalkFoldersFail:
    BugMessage Err.Description
End Function

' Display a context menu from a folder
' Based on C code by Jeff Procise in PC Magazine
' Destroys any pidl passed to it, so pass duplicate if necessary
Function ContextPopMenu(ByVal hWnd As Long, vItem As Variant, _
                        ByVal x As Long, ByVal y As Long) As Boolean
    InitIf  ' Initialize if in standard modue

    Dim folder As IVBShellFolder, pidlMenu As Long
    Dim menu As IVBContextMenu, ici As CMINVOKECOMMANDINFO
    Dim iCmd As Long, f As Boolean, hMenu As Long
    
    ' Get folder and pidl from path, pidl, or special item
    Set folder = FolderFromItem(vItem, pidlMenu)
    If folder Is Nothing Then Exit Function

    ' Get an IContextMenu object
    On Error GoTo ContextPopMenuFail
    folder.GetUIObjectOf hWnd, 1, pidlMenu, iidContextMenu, 0, menu

    ' Create an empty popup menu and initialize it with QueryContextMenu
    hMenu = CreatePopupMenu
    On Error GoTo ContextPopMenuFail2
    menu.QueryContextMenu hMenu, 0, 1, &H7FFF, CMF_EXPLORE
        
    ' Convert x and y to client coordinates
    MWinTool.ClientToScreenXY hWnd, x, y
    
    ' Display the context menu
    Const afMenu = TPM_LEFTALIGN Or TPM_LEFTBUTTON Or _
                   TPM_RIGHTBUTTON Or TPM_RETURNCMD
    iCmd = TrackPopupMenu(hMenu, afMenu, x, y, 0, hWnd, ByVal hNull)

    ' If a command was selected from the menu, execute it.
    If iCmd Then
        ici.cbSize = LenB(ici)
        ici.fMask = 0
        ici.hWnd = hWnd
        ici.lpVerb = iCmd - 1
        ici.lpParameters = pNull
        ici.lpDirectory = pNull
        ici.nShow = SW_SHOWNORMAL
        ici.dwHotKey = 0
        ici.hIcon = hNull
        menu.InvokeCommand ici
        ContextPopMenu = True
    End If
    
ContextPopMenuFail2:
    DestroyMenu hMenu
           
ContextPopMenuFail:
    ' Menu pidl is freed, so client had better not pass only copy
    Allocator.Free pidlMenu
    BugMessage Err.Description

End Function

' Recent document list

Sub AddToRecentDocs(sDoc As String)
    SHAddToRecentDocs SHARD_PATH, sDoc
End Sub

Sub ClearRecentDocs()
    SHAddToRecentDocs SHARD_PATH, sNullStr
End Sub

' Thanks to Magnus Lindström for suggesting default directory fix
Function BrowseForFolder(Optional Owner As Long = hNull, _
                         Optional DisplayName As String, _
                         Optional Options As EBROWSEOPTIONS, _
                         Optional Title As String, _
                         Optional Root As Variant, _
                         Optional Default As String) As String
    InitIf  ' Initialize if in standard modue
    
    Dim bi As BROWSEINFO, sOut As String, sTitle As String
    bi.hwndOwner = Owner
    sOut = String$(cMaxPath, 0)
    bi.pszDisplayName = MUtility.StringToPointer(sOut)
    If UnicodeTypeLib Then
        sTitle = Title
    Else
        sTitle = StrConv(Title, vbFromUnicode)
    End If
    bi.lpszTitle = MUtility.StringToPointer(sTitle)

    ' BIF_BROWSEFORCOMPUTER     Only return computers
    ' BIF_BROWSEFORPRINTER      Only return printers
    ' BIF_BROWSEINCLUDEFILES    Display files as well as folders
    ' BIF_DONTGOBELOWDOMAIN     Do not include network folders below the domain level
    ' BIF_EDITBOX               Include edit control for typing item names
    ' BIF_RETURNFSANCESTORS     Only return file system ancestors
    ' BIF_RETURNONLYFSDIRS      Only return file system directories
    ' BIF_STATUSTEXT            Include a status area
    ' BIF_VALIDATE              Refuse invalid names by sending BFFM_VALIDATEFAILED
    bi.ulFlags = Options 'Or BIF_EDITBOX Or BIF_VALIDATE Or BIF_STATUSTEXT
    ' bi.iImage = 0
    Dim pidlIn As Long, pidlOut As Long, sPath As String
    
    If IsMissing(Root) Then
        pidlIn = PidlFromSpecialFolder(CSIDL_DRIVES, Owner)
    ElseIf VarType(Root) = vbString Then
        ' Start specified as string path
        sPath = MUtility.NormalizePath(CStr(Root))
        pidlIn = PathToPidl(sPath)
    Else
        ' Start specified as pidl
        pidlIn = ToPidl(Root)
    End If
    bi.pidlRoot = pidlIn
    
    Dim sDefault As String
    If Default <> sEmpty Then
        sDefault = MUtility.NormalizePath(Default)
        If UnicodeTypeLib = False Then
            sDefault = StrConv(sDefault, vbFromUnicode)
        End If
        bi.lParam = MUtility.StringToPointer(sDefault)
    End If
    bi.lpfn = MUtility.GetProc(AddressOf BrowseCallbackProc)
    
    pidlOut = SHBrowseForFolder(bi)
    DisplayName = MUtility.PointerToString(bi.pszDisplayName)
    BrowseForFolder = PathFromPidl(pidlOut)
    
    ' Free the pidls we create
    If IsMissing(Root) Then
        Allocator.Free pidlIn
    ElseIf VarType(Root) = vbString Then
        Allocator.Free pidlIn
    Else
        ' Leave the pidl we received as a parameter
    End If

End Function

#If fComponent = 0 Then
' This function for standard module only--global module version
' must be in separate file
Function BrowseCallbackProc(ByVal hWnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal lParam As Long, _
                            ByVal lpData As Long) As Long
    Dim s As String, iRet As Long
    Select Case uMsg
    ' Browse dialog box has finished initializing (lParam is NULL)
    Case BFFM_INITIALIZED
        Debug.Print "BFFM_INITIALIZED"
        s = MUtility.PointerToString(lpData)
        MUtility.DenormalizePath s
        ' Set the selection
        iRet = SendMessageStr(hWnd, BFFM_SETSELECTION, ByVal APITRUE, s)
        
    ' Selection has changed (lParam contains pidl of selected folder)
    Case BFFM_SELCHANGED
        Debug.Print "BFFM_SELCHANGED"
        ' Display full path if status area if enabled
        s = MFoldTool.PathFromPidl(lParam)
        iRet = SendMessageStr(hWnd, BFFM_SETSTATUSTEXT, ByVal 0&, s)
        
    ' Invalid name in edit box (lParam parameter has invalid name string)
    Case BFFM_VALIDATEFAILED
        Debug.Print "BFFM_VALIDATEFAILED"
        ' Return zero to dismiss dialog or nonzero to keep it displayed
        ' Disable the OK button
        iRet = SendMessage(hWnd, BFFM_ENABLEOK, ByVal 0&, ByVal APIFALSE)
        s = MUtility.PointerToString(lParam)
        s = "Path invalid: " & s
        iRet = SendMessageStr(hWnd, BFFM_SETSTATUSTEXT, ByVal 0&, s)

    Case Else
        Debug.Print uMsg
    End Select
    BrowseCallbackProc = 0
End Function

Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".FoldTool"
        Select Case e
        Case eeBaseFoldTool
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

