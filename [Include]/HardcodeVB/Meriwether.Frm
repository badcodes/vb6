VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FMeriwether 
   Caption         =   "Meriwether"
   ClientHeight    =   5736
   ClientLeft      =   636
   ClientTop       =   3456
   ClientWidth     =   8964
   ClipControls    =   0   'False
   Icon            =   "Meriwether.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Meriwether.frx":0442
   ScaleHeight     =   5736
   ScaleWidth      =   8964
   Begin MSComctlLib.Toolbar bar 
      Align           =   1  'Align Top
      Height          =   312
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8964
      _ExtentX        =   15812
      _ExtentY        =   550
      ButtonWidth     =   487
      ButtonHeight    =   466
      ImageList       =   "imlstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2200
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up"
            Description     =   "Up One Level"
            Object.ToolTipText     =   "Up One Level"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Map"
            Description     =   "Map Network Drive"
            Object.ToolTipText     =   "Map Network Drive"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Unmap"
            Description     =   "Disconnect Network Drive"
            Object.ToolTipText     =   "Disconnect Network Drive"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Description     =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Description     =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Description     =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Large"
            Description     =   "Large Icons"
            Object.ToolTipText     =   "Large Icons"
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Small"
            Description     =   "Small Icons"
            Object.ToolTipText     =   "Small Icons"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "List"
            Description     =   "List"
            Object.ToolTipText     =   "List"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Details"
            Description     =   "Details"
            Object.ToolTipText     =   "Details"
            ImageIndex      =   13
            Style           =   2
         EndProperty
      EndProperty
      Begin VB.CommandButton cmdDir 
         Caption         =   "..."
         Height          =   270
         Left            =   1845
         TabIndex        =   5
         Top             =   24
         Width           =   270
      End
      Begin VB.TextBox txtDir 
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   45
         Width           =   1656
      End
   End
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   348
      Left            =   0
      TabIndex        =   2
      Top             =   5388
      Width           =   8964
      _ExtentX        =   15812
      _ExtentY        =   614
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Objects"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12848
            Text            =   "Add your status text here"
            TextSave        =   "Add your status text here"
            Key             =   "Status"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   4344
      Left            =   2856
      TabIndex        =   1
      Top             =   960
      Width           =   6228
      _ExtentX        =   10986
      _ExtentY        =   7662
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwFiles 
      Height          =   4848
      Left            =   96
      TabIndex        =   0
      Top             =   468
      Width           =   2712
      _ExtentX        =   4784
      _ExtentY        =   8551
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlstSIcon 
      Left            =   7080
      Top             =   450
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlstLIcon 
      Left            =   6150
      Top             =   435
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlstToolbar 
      Left            =   8025
      Top             =   435
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":0594
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":0736
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":08D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":0A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":0C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":0DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":0F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":1102
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":12A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":1446
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":15E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":178A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Meriwether.frx":192C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Large Icons"
         Index           =   0
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Small Icons"
         Index           =   1
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Details"
         Index           =   3
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnyArrange 
         Caption         =   "&Arrange"
         Begin VB.Menu mnuBy 
            Caption         =   "by &Name"
            Index           =   0
         End
         Begin VB.Menu mnuBy 
            Caption         =   "by &Type"
            Index           =   1
         End
         Begin VB.Menu mnuBy 
            Caption         =   "by Si&ze"
            Index           =   2
         End
         Begin VB.Menu mnuBy 
            Caption         =   "by &Date"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "FMeriwether"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IUseFolder

' Private variables
Private split As New CSplitter
Private xList As Single, yList As Single
Private fWalkNodes As Boolean
Private fFirstItem As Boolean
Private sCurNode As String

' Names of special folder locations
Private sDrives As String
Private sBitBucket As String
Private sNetwork As String
Private sFonts As String
Private sPrinters As String
Private sControlPanel As String


Private Sub Form_Load()
    
    Dim file As New CFileInfo
    If Not HasShell Then
        MsgBox "This program requires the Windows enhanced user interface"
        End
    End If
    ' Set up splitter
    split.Create LeftControl:=tvwFiles, _
                 RightControl:=lvwFiles, _
                 Vertical:=True, _
                 BorderPixels:=4, _
                 AutoBorder:=True, _
                 Resizeable:=True, _
                 Percent:=35
    SetDisplay lvwReport, True
    Show
    
    ' Initialize special folder names
    sDrives = NameFromPidl(ToPidl(CSIDL_DRIVES))
    sBitBucket = NameFromPidl(ToPidl(CSIDL_BITBUCKET))
    sNetwork = NameFromPidl(ToPidl(CSIDL_NETWORK))
    sFonts = NameFromPidl(ToPidl(CSIDL_FONTS))
    sPrinters = NameFromPidl(ToPidl(CSIDL_PRINTERS))
    sControlPanel = NameFromPidl(ToPidl(CSIDL_CONTROLS))

    Refresh
    HourGlass Me
    Dim nodeRoot As Node
    Set nodeRoot = InitDesktop
    Refresh
    nodeRoot.Child.EnsureVisible
    Refresh
    tvwFiles_NodeClick nodeRoot
    HourGlass Me
    bar.Refresh
    
End Sub

' Special case node initialization for root
Function InitDesktop() As Node
    Dim folder As IVBShellFolder, pidl As Long, sKey As String
    ' Get folder and pidl roots
    Set folder = GetDesktopFolder
    pidl = ToPidl(CSIDL_DESKTOP)
    ' Create large and small image lists
    With FileInfoFromFolder(folder, pidl)
        imlstLIcon.ListImages.Add , .DisplayName, .LargeIcon()
        imlstSIcon.ListImages.Add , .DisplayName, .SmallIcon()
        ' Set TreeView and ListView to use image lists
        lvwFiles.Icons = imlstLIcon
        lvwFiles.SmallIcons = imlstSIcon
        tvwFiles.ImageList = imlstSIcon
        ' Add the desktop node
        Set InitDesktop = tvwFiles.Nodes.Add(, , .DisplayName, _
                                             .DisplayName, .DisplayName)
        ' Mark as having children
        Call tvwFiles.Nodes.Add(InitDesktop.Index, tvwChild, , ":")
        ' Save subfolder in tag
        InitDesktop.Tag = folder
    End With
End Function

Sub InitNode(folder As IVBShellFolder, ByVal pidl As Long, _
             nodeCur As Variant)
    ' Create file information object from folder and pidl
    Dim fi As New CFileInfo, nodeNew As Node
    Dim sIconKey As String, sKey As String
    sCurNode = GetFolderName(folder, pidl, SHGDN_FORPARSING)
    Set fi = FileInfoFromFolder(folder, pidl)
    With fi
        sIconKey = .DisplayName
        If .ItemType = eitFile And .TypeName <> "Shortcut" Then
            sIconKey = .TypeName
        End If
        ' Add icons
        AddIconImages imlstLIcon, imlstSIcon, fi, sIconKey
        ' Add new node
        If .ItemType = eitFile Or .ItemType = eitDrive Then
            sKey = GetFolderName(folder, pidl, SHGDN_FORPARSING)
        Else
            sKey = .DisplayName
        End If
        Set nodeNew = tvwFiles.Nodes.Add(nodeCur.Index, tvwChild, sKey, _
                                         .DisplayName, sIconKey)
    End With
    
    ' Check for children
    Dim afAttr As ESFGAO
    afAttr = SFGAO_FOLDER Or SFGAO_HASSUBFOLDER
    folder.GetAttributesOf 1, pidl, afAttr
    If afAttr And (SFGAO_FOLDER Or SFGAO_HASSUBFOLDER) Then
        ' Mark as having children
        Call tvwFiles.Nodes.Add(nodeNew.Index, tvwChild, , ":")
        ' Get subfolder and save it in tag
        Dim folderNew As IVBShellFolder
        Set folderNew = BindToShell(folder, pidl)
        nodeNew.Tag = folderNew
    End If
End Sub

Sub InitItem(folder As IVBShellFolder, ByVal pidl As Long, _
             nodeCur As Variant)
    ' Create file information object from folder and pidl
    Dim fi As New CFileInfo, nodeNew As Node
    Dim sName As String, sKey As String, Item As ListItem
    If fFirstItem Then InitColumn fi, nodeCur.Text
    Set fi = FileInfoFromFolder(folder, pidl)
    With fi
        sName = .DisplayName
        sKey = sName
        If .ItemType = eitFile And .TypeName <> "Shortcut" Then
            sKey = .TypeName
        End If
        ' Add icons
        AddIconImages imlstLIcon, imlstSIcon, fi, sKey
        ' Add list item
        Set Item = lvwFiles.ListItems.Add(, , sName, sKey, sKey)
        ' Add subitems if appropriate
        Select Case .ItemType
        Case eitFile
            If .Length Then Item.SubItems(1) = CStr(.Length \ 1000) & " KB"
            Item.SubItems(2) = .TypeName
            Item.SubItems(3) = Format$(.Modified, "mm/dd/yy hh:mm AMPM")
            Item.SubItems(4) = Format$(.Created, "mm/dd/yy hh:mm AMPM")
            Item.SubItems(5) = Format$(.Accessed, "mm/dd/yy")
        Case eitDrive
            Item.SubItems(1) = .DriveType
            Item.SubItems(2) = Format$(.TotalKilo, "##,##0") & "MB"
            Item.SubItems(3) = Format$(.FreeKilo, "##,##0") & "MB"
        ' Case Else
            ' No subitems
        End Select
    End With

End Sub

Sub InitColumn(fi As CFileInfo, sNodeText As String)
    Dim col As ColumnHeader, dx As Single
    lvwFiles.ColumnHeaders.Clear
    ' Create ListView ColumnHeader
    Select Case sNodeText
    Case sBitBucket, sNetwork, sFonts, sPrinters, sControlPanel
        ' Special folders with non-file items
        dx = TextWidth("Very, Very Long File Name")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Name", dx)
        col.Alignment = sbrLeft
    Case sDrives ' My Computer by default, but could be changed
        ' Contains mostly drives
        dx = TextWidth("Root on 'LabelName' (D:)")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Name", dx)
        col.Alignment = sbrLeft
        dx = TextWidth("Floppy Missing")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Type", dx)
        col.Alignment = sbrLeft
        dx = TextWidth("Total Size")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Total Size", dx)
        col.Alignment = sbrRight
        dx = TextWidth("Free Space")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Free Space", dx)
        col.Alignment = sbrRight
    Case Else
        ' Set up sizes for files
        dx = TextWidth("Very Long File Name")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Name", dx)
        col.Alignment = sbrLeft
        dx = TextWidth("9,999 KB")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Size", dx)
        col.Alignment = sbrRight
        dx = TextWidth("MS-DOS Batch File")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Type", dx)
        col.Alignment = sbrLeft
        dx = TextWidth("12/25/97 12:01 AM")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Modified", dx)
        col.Alignment = sbrLeft
        Set col = lvwFiles.ColumnHeaders.Add(, , "Created", dx)
        col.Alignment = sbrLeft
        dx = TextWidth("12/25/97")
        Set col = lvwFiles.ColumnHeaders.Add(, , "Accessed", dx)
        col.Alignment = sbrLeft
    End Select
    fFirstItem = False
End Sub

Private Function IUseFolder_UseFolder(UserData As Variant, _
                                      CurFolder As Win.IVBShellFolder, _
                                      ByVal ItemList As Long) As Boolean
    ' Display a node or item
    If fWalkNodes Then
        InitNode CurFolder, ItemList, UserData
    Else
        InitItem CurFolder, ItemList, UserData
    End If
End Function

Private Sub cmdDir_Click()
    Dim s As String, af As Long
    af = BIF_STATUSTEXT
    s = BrowseForFolder(hWnd, , af, "Explore to the sea", 0, txtDir.Text)
    ' To do: Should find and set path in TreeView
    If s <> sEmpty Then
        txtDir.Text = s
        MsgBox "Find this directory in the TreeView: " & txtDir.Text
    End If
End Sub

Private Sub tvwFiles_Expand(ByVal Node As Node)
    BugMessage "Expand: " & Node.Text
    ' Remove fake child and insert real one
    BugAssert Node.Child.Text = ":"
    tvwFiles.Nodes.Remove Node.Child.Index
    ' Processs subfolders recursively
    Dim folder As IVBShellFolder, pidl As Long
    Set folder = Node.Tag
    fWalkNodes = True
    
    WalkFolders folder, Me, Node, ewmFolders

End Sub

Private Sub tvwFiles_Collapse(ByVal Node As Node)
    BugMessage "Collapse: " & Node.Text
    Dim i As Integer
    ' Remove real children and replace with fake child
    With Node.Child
        For i = .LastSibling.Index To .Index Step -1
            tvwFiles.Nodes.Remove i
        Next
    End With
    tvwFiles.Nodes.Add Node.Index, tvwChild, , ":"
End Sub

Private Sub tvwFiles_NodeClick(ByVal Node As Node)
    BugMessage "NodeClick: " & Node.Text
    ' Process items
    lvwFiles.ListItems.Clear
    fWalkNodes = False
    fFirstItem = True
    txtDir.Text = Node.Key
    WalkFolders Node.Tag, Me, Node, ewmNonfolders
    
    With stat.Panels("Objects")
        Select Case lvwFiles.ListItems.Count
        Case 0
            .Text = sEmpty
        Case 1
            .Text = "1 object"
        Case Else
            .Text = lvwFiles.ListItems.Count & " objects"
        End Select
    End With
    ' To do: Add more status text
End Sub

Private Sub lvwFiles_ItemClick(ByVal Item As ListItem)
    ' Challenge: Put information in status bar
    BugMessage "Clicked item: " & Item.Text
End Sub

' ListView doesn't have ItemDblClick so fake it
Private Sub lvwFiles_ItemDblClick(ByVal Item As ListItem)
    If Item Is Nothing Then Exit Sub
    Dim s As String
    On Error Resume Next
    s = Item.SubItems(2)
    If s = "File Folder" Then
        BugMessage "Double Clicked Directory: " & Item.Text
        ' Challenge: Need to trace through file path to find
        ' and open appropriate directory in TreeView
    Else
        BugMessage "Double clicked item: " & Item.Text
        ' To do: Turn Item.Text into a path so we can execute it
        VBShellExecute Item.Text
    End If
End Sub

' Save x and y for testing in DblClick
Private Sub lvwFiles_MouseDown(Button As Integer, Shift As Integer, _
                               x As Single, y As Single)
    xList = x
    yList = y
End Sub

Private Sub lvwFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim Item As ListItem, sPath As String
        Set Item = lvwFiles.HitTest(xList, yList)
        If Mid$(sCurNode, 2, 2) = ":\" Then
            sPath = sCurNode & "\" & Item.Text
        Else
            sPath = Item.Text
        End If
        ' To do: Make right click work
        Dim f As Boolean
        f = ContextPopMenu(lvwFiles.hWnd, sPath, xList, yList)
    End If
End Sub

' Create missing ItemDblClick event
Private Sub lvwFiles_DblClick()
    lvwFiles_ItemDblClick lvwFiles.HitTest(xList, yList)
End Sub

Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As ColumnHeader)
With lvwFiles
    Static iLast As Integer, iCur As Integer
    .Sorted = True
    iCur = ColumnHeader.Index - 1
    If iCur = iLast Then .SortOrder = IIf(.SortOrder = 1, 0, 1)
    lvwFiles.SortKey = iCur
    iLast = iCur
End With
End Sub

Private Sub bar_ButtonClick(ByVal Button As Button)
    Select Case Button.Key
    Case "Large"
        SetDisplay lvwIcon, False
    Case "Small"
        SetDisplay lvwSmallIcon, False
    Case "List"
        SetDisplay lvwList, False
    Case "Details"
        SetDisplay lvwReport, False
    Case Else
        MsgBox "Left as an exercise for the reader"
    End Select
End Sub

Sub AddIconImages(imlstLIcon As ImageList, imlstSIcon As ImageList, _
                  fi As CFileInfo, sKey As String)
    Dim i As Integer
    On Error Resume Next
    i = imlstLIcon.ListImages(sKey).Index
    If Err Then
        On Error GoTo 0
        imlstLIcon.ListImages.Add , sKey, fi.LargeIcon()
        imlstSIcon.ListImages.Add , sKey, fi.SmallIcon()
    ' Else
        ' Already added
    End If
End Sub

Sub SetDisplay(ByVal Index As Integer, fMenu As Boolean)
    Dim i As Integer
    ' Check menu item
    For i = 0 To 3
        mnuDisplay(i).Checked = (i = Index)
    Next
    lvwFiles.View = Index
    ' Push toolbar button
    If fMenu Then
        i = bar.Buttons("Large").Index + Index
        bar.Buttons(i).Value = tbrPressed
    End If
End Sub

Private Sub mnuBy_Click(Index As Integer)
    lvwFiles.Sorted = True
    lvwFiles.SortKey = Index
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuDisplay_Click(Index As Integer)
    SetDisplay Index, True
End Sub

Private Sub mnuEdit_Click()
    MsgBox "Left as an exercise for the reader"
End Sub

Private Sub mnuHelp_Click()
    MsgBox "Left as an exercise for the reader"
End Sub

Private Sub mnuNew_Click()
    MsgBox "Left as an exercise for the reader"
End Sub

Private Sub mnuStatus_Click()
    MsgBox "Left as an exercise for the reader"
End Sub

Private Sub mnuTools_Click()
    MsgBox "Left as an exercise for the reader"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    split.Splitter_MouseDown Button, Shift, x, y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    split.Splitter_MouseMove Button, Shift, x, y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    split.Splitter_MouseUp Button, Shift, x, y
End Sub

Private Sub Form_Resize()
    split.Splitter_Resize
End Sub

