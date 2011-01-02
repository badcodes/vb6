VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "TemplateClassBuilder"
   ClientHeight    =   9555
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10605
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Quick Process"
      Height          =   435
      Index           =   2
      Left            =   150
      TabIndex        =   19
      Top             =   8910
      Width           =   3840
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&Guess"
      Height          =   375
      Index           =   4
      Left            =   9090
      TabIndex        =   18
      Top             =   3780
      Width           =   1395
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Start Process"
      Height          =   435
      Left            =   6840
      TabIndex        =   17
      Top             =   8940
      Width           =   2235
   End
   Begin VB.TextBox txtSource 
      Appearance      =   0  'Flat
      Height          =   3795
      HideSelection   =   0   'False
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   4860
      Width           =   10095
   End
   Begin VB.TextBox txtTarget 
      Appearance      =   0  'Flat
      Height          =   3795
      HideSelection   =   0   'False
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      Top             =   4860
      Width           =   10095
   End
   Begin MSComctlLib.TabStrip tabFile 
      Height          =   4275
      Left            =   120
      TabIndex        =   14
      Top             =   4500
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7541
      TabWidthStyle   =   2
      TabFixedWidth   =   7056
      TabFixedHeight  =   459
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Source"
            Key             =   "source"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Target"
            Key             =   "target"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Save"
      Height          =   435
      Index           =   1
      Left            =   9240
      TabIndex        =   13
      Top             =   8940
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
      Height          =   435
      Index           =   0
      Left            =   5460
      TabIndex        =   12
      Top             =   8940
      Width           =   1215
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&Clear"
      Height          =   375
      Index           =   3
      Left            =   9090
      TabIndex        =   11
      Top             =   3270
      Width           =   1395
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&Delete"
      Height          =   375
      Index           =   2
      Left            =   9075
      TabIndex        =   10
      Top             =   2790
      Width           =   1395
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&Edit"
      Height          =   375
      Index           =   1
      Left            =   9075
      TabIndex        =   9
      Top             =   2340
      Width           =   1395
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&Add"
      Height          =   375
      Index           =   0
      Left            =   9075
      TabIndex        =   8
      Top             =   1890
      Width           =   1395
   End
   Begin MSComctlLib.ListView typeList 
      Height          =   2295
      Left            =   240
      TabIndex        =   7
      Top             =   1980
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "id"
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "from"
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "to"
         Text            =   "To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "style"
         Text            =   "Style"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame frmOption 
      Caption         =   "Type Table"
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   8775
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Save&To..."
      Height          =   375
      Index           =   1
      Left            =   9060
      TabIndex        =   4
      Top             =   1200
      Width           =   1395
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1260
      Width           =   8775
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Open..."
      Height          =   375
      Index           =   0
      Left            =   9060
      TabIndex        =   1
      Top             =   420
      Width           =   1395
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   8775
   End
   Begin VB.Label lblAnything 
      Caption         =   "Select Target File"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   900
      Width           =   12495
   End
   Begin VB.Label lblAnything 
      Caption         =   "Select Template Source:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   12495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const FILENAME_FILTER As String = "VB Source File|*.bas;*.cls;*.frm;*.vbp|All Files(*.*)|*.*"
Private mTypeString(0 To 2) As String

Private Sub cmdAction_Click(Index As Integer)
    If Index = 2 Then
        If txtFilename(0).Text <> "" Then
            Dim pDir As String
            Dim pName As String
            pDir = MFileSystem.GetParentFolderName(txtFilename(0).Text)
            pName = MFileSystem.GetFileName(txtFilename(0).Text)
            pDir = BuildPath(pDir)
            BuildClass pName, "", pDir
        End If
        Exit Sub
    End If
    If txtFilename(1).Text <> "" Then
        SaveFilename txtFilename(1).Text
    Else
        cmdSelect_Click 1
        If txtFilename(1).Text <> "" Then SaveFilename txtFilename(1).Text
    End If
End Sub

Private Sub cmdList_Click(Index As Integer)
    Select Case Index
        Case 0  'Add
            AddEmptyItem
        Case 1
            EditListItem (typeList.SelectedItem.Index)
        Case 2
            DeleteListItem (typeList.SelectedItem.Index)
        Case 3
            ClearList
        Case 4
            AddTypeItemByFilename
    End Select
End Sub

Private Sub AddTypeItemByFilename()
    Dim pSrc As String
    pSrc = txtFilename(0).Text
    Dim pDst As String
    pDst = txtFilename(1).Text
    
    Dim pSrcName As String
    Dim pDstName As String
    pSrcName = MFileSystem.GetBaseName(pSrc)
    pDstName = MFileSystem.GetBaseName(pDst)
    
    If StrComp(Left(pSrcName, 1), "T", vbTextCompare) = 0 Then
        pSrcName = Mid$(pSrcName, 2)
    End If
    
    If StrComp(pSrcName, Right$(pDstName, Len(pSrcName))) = 0 Then
        Dim pTypeA As String
        pTypeA = Left$(pDstName, Len(pDstName) - Len(pSrcName))
        If pTypeA = "" Then pTypeA = "Variant"
        Dim pTypeAType As String
        Select Case LCase$(pTypeA)
            Case "", "variant"
                pTypeAType = mTypeString(2)
            Case "string", "long", "integer", "boolean", "byte"
                pTypeAType = mTypeString(0)
            Case Else
                pTypeAType = mTypeString(1)
        End Select
        Dim item As ListItem
        Set item = typeList.ListItems.Add(, , "#")
        item.ListSubItems.Add , , "TPLAType"
        item.ListSubItems.Add , , pTypeA
        item.ListSubItems.Add , , pTypeAType
        Set item = typeList.ListItems.Add(, , "#")
        item.ListSubItems.Add , , "T" & pSrcName
        item.ListSubItems.Add , , "C" & pDstName
        item.ListSubItems.Add , , "Object"
    End If
    
    
End Sub

'CSEH: ErrorAbort
Private Sub cmdProcess_Click()
    '<EhHeader>
    On Error GoTo cmdProcess_Click_Abort

    '</EhHeader>
   tabFile.Tabs("target").Selected = True
    txtTarget.Text = ""
    
    Dim fnTempIn As String
    Dim fnTempOut As String
    
    fnTempIn = GetTempFilename
    fnTempOut = GetTempFilename
    FS_WriteFile fnTempIn, txtSource.Text
    
    Dim Reportor As TextBoxReportor
    Set Reportor = New TextBoxReportor
    
    Dim processor As CTemplateBuilder
    Set processor = New CTemplateBuilder
    
    Dim typeStyle As CTypeStyle
    Set typeStyle = New CTypeStyle
    With typeStyle
        .ConstVarOf(CTTypeNormal) = "NormalType"
        .ConstVarOf(CTTypeObject) = "ObjectType"
        .ConstVarOf(CTTypeVariant) = "VariantType"
    End With
       
    Dim templateType As CType
    Set templateType = New CType
    Set templateType.typeStyle = typeStyle
    
    templateType.ConstTypePrefix = "f"
    templateType.ConstTypeSuffix = ""
    
    Dim i As Long
    With processor
        .InitType templateType
        .AddFilter New CFilterModule
        .AddFilter New CFilterConstVar
        .AddFilter New CFilterTypeName
        .AddFilter New CFilterTypeOP
        For i = 1 To typeList.ListItems.count
            .AddType typeList.ListItems(i).SubItems(1), _
                     typeList.ListItems(i).SubItems(2), _
                     TypeStyleFrom(typeList.ListItems(i).SubItems(3))
        Next
        Set .Reportor = Reportor
    End With
    If processor.Process(fnTempIn, fnTempOut) Then
        LoadFilename fnTempOut, 1
        If FileExists(fnTempIn) Then Kill fnTempIn
        If FileExists(fnTempOut) Then Kill fnTempOut
    Else
        MsgBox "Error when processing", vbCritical
    End If
    
    '<EhFooter>
    Exit Sub

cmdProcess_Click_Abort:
    On Error Resume Next
    Debug.Print Err.Number; ":" & Err.Description
    If FileExists(fnTempIn) Then Kill fnTempIn
    If FileExists(fnTempOut) Then Kill fnTempOut
    '</EhFooter>
End Sub
Private Function TypeStyleFrom(ByVal sType As String) As CTTypeStyles
    Dim i As Long
    Dim idx As Long
    For i = LBound(mTypeString) To UBound(mTypeString)
        If sType = mTypeString(i) Then idx = i: Exit For
    Next
    If idx = 0 Then
        TypeStyleFrom = CTTypeNormal
    ElseIf idx = 1 Then
        TypeStyleFrom = CTTypeObject
    Else
        TypeStyleFrom = CTTypeVariant
    End If
    
End Function
Private Sub cmdSelect_Click(Index As Integer)
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    Dim FileName As String
    If Index = 0 Then
        FileName = txtFilename(0).Text
        If dlg.VBGetOpenFileName(FileName, , , , , , FILENAME_FILTER) Then
            txtFilename(0).Text = FileName
            LoadFilename FileName, 0
        End If
    ElseIf Index = 1 Then
        FileName = txtFilename(1).Text
        If FileName = "" Then FileName = txtFilename(0).Text
        If dlg.VBGetSaveFileName(FileName, , , FILENAME_FILTER) Then
            txtFilename(1).Text = FileName
        End If
    End If
    Set dlg = Nothing
End Sub

Private Sub Form_Load()
    mTypeString(0) = "Normal"
    mTypeString(1) = "Object"
    mTypeString(2) = "Variant"
    Form_Resize
    
    Dim mySetting As CVBSetting
    Set mySetting = New CVBSetting
    With mySetting
        .Appname = App.EXEName
        .Section = "MainForm"
        
        .ReadPropListItems typeList, "TypeList"
        .ReadPropText txtFilename(0), "Source", ""
        .ReadPropText txtFilename(1), "Target", ""
    End With
    Set mySetting = Nothing
    If txtFilename(0).Text <> "" Then LoadFilename txtFilename(0).Text, 0
End Sub

Private Sub Form_Resize()
    With typeList
        Dim nCol As Integer
        Dim lCol As Double
        nCol = .ColumnHeaders.count
        lCol = (.Width - .ColumnHeaders(1).Width) / (nCol - 1)
        Do Until nCol < 2
            .ColumnHeaders(nCol).Width = lCol
            nCol = nCol - 1
        Loop
    End With
    With txtSource
        txtTarget.Move .Left, .tOP, .Width, .Height
    End With
End Sub

Private Sub AddEmptyItem()
    Dim item As ListItem
    Set item = typeList.ListItems.Add(, , "#")
    item.ListSubItems.Add , , ""
    item.ListSubItems.Add , , ""
    item.ListSubItems.Add , , ""
    EditListItem item.Index
End Sub




Private Sub Form_Unload(Cancel As Integer)
    Dim mySetting As CVBSetting
    Set mySetting = New CVBSetting
    With mySetting
        .Appname = App.EXEName
        .Section = "MainForm"
        .WritePropText txtFilename(0), "Source"
        .WritePropText txtFilename(1), "Target"
        .WritePropListItems typeList, "TypeList"
    End With
    Set mySetting = Nothing
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub tabFile_Click()
    If tabFile.SelectedItem.Index = 1 Then
        txtSource.ZOrder 0
    Else
        txtTarget.ZOrder 0
    End If
End Sub


Private Sub EditListItem(ByRef idx As Long)
    On Error Resume Next
    If idx > typeList.ListItems.count Or idx < 1 Then Exit Sub
    Dim EditList As dlgEditList
    Set EditList = New dlgEditList
    
    Load EditList
    EditList.Icon = Me.Icon
    EditList.IsOK = 0
    With EditList.TextField(0)
        Dim i As Long
        For i = 1 To 26
            .AddItem MClassTemplate.GetTypeName(i)
        Next
    End With
    With EditList.TextField(1)
        .AddItem "String"
        .AddItem "Long"
        .AddItem "Integer"
        .AddItem "Variant"
        .AddItem "Byte"
    End With
    For i = LBound(mTypeString) To UBound(mTypeString)
        EditList.TextField(2).AddItem mTypeString(i)
    Next
    With typeList.ListItems(idx)
        EditList.TextField(0).Text = .SubItems(1)
        EditList.TextField(1).Text = .SubItems(2)
        EditList.TextField(2).Text = IIf(.SubItems(3) <> "", .SubItems(3), mTypeString(0))
    End With
    EditList.Show 1
    If EditList.IsOK > 0 Then
        With typeList.ListItems(idx)
            .SubItems(1) = EditList.TextField(0).Text
            .SubItems(2) = EditList.TextField(1).Text
            .SubItems(3) = EditList.TextField(2).Text
        End With
    End If
    'Unload EditList
    Unload EditList
    Set EditList = Nothing
    Debug.Print EditList Is Nothing
End Sub

Private Sub DeleteListItem(ByRef idx As Long)
    If idx > typeList.ListItems.count Or idx < 1 Then Exit Sub
    typeList.ListItems.Remove (idx)
End Sub

Private Sub ClearList()
    typeList.ListItems.Clear
End Sub

'CSEH: ErrorAbort
Private Sub LoadFilename(ByRef sFilename As String, ByVal idx As Integer)
    '<EhHeader>
    On Error GoTo LoadFilename_Abort

    '</EhHeader>
    Dim n As Integer
    n = FreeFile
    Open sFilename For Input As #n
    If (idx) = 0 Then
        txtSource.Text = Input(LOF(n), n)
    ElseIf idx = 1 Then
        txtTarget.Text = Input(LOF(n), n)
    End If
    Close #n
    
    Exit Sub

LoadFilename_Abort:
    MsgBox Err.Description, vbCritical
    On Error Resume Next
    Close #n
    '</EhFooter>
End Sub

Private Sub SaveFilename(ByRef sFilename As String)
    Dim n As Integer
    n = FreeFile
    Open sFilename For Output As #n
    Print #n, txtTarget.Text;
    Close #n
    MsgBox "Ok,File saved to " & sFilename
    Exit Sub
File_Error:
    On Error Resume Next
    MsgBox Err.Description, vbCritical
    Err.Clear
    Close #n
End Sub

Private Sub txtFilename_Change(Index As Integer)
    typeList.ListItems.Clear
    AddTypeItemByFilename
End Sub

Private Sub typeList_DblClick()
    EditListItem typeList.SelectedItem.Index
End Sub

Public Sub BuildClass(ByVal vsrcName As String, ByVal vdstname As String, _
    Optional ByVal vDir As String = "X:\Workspace\VB6\[Include]\ClassTemplate\")
Debug.Print vsrcName, vdstname
If vdstname = "" Then
    Dim basename As String
    basename = Mid$(vsrcName, 2)
    BuildClass vsrcName, basename
    BuildClass vsrcName, "Object" & basename
    BuildClass vsrcName, "String" & basename
    BuildClass vsrcName, "Long" & basename
    BuildClass vsrcName, "Integer" & basename
    BuildClass vsrcName, "Byte" & basename
    BuildClass vsrcName, "Boolean" & basename
    Exit Sub
End If
    Dim pSrc As String
    pSrc = vDir & vsrcName
    Dim pDst As String
    pDst = vDir & vdstname
    txtFilename(0).Text = pSrc
    txtFilename(1).Text = pDst
    cmdProcess_Click
    cmdAction_Click 1
End Sub

