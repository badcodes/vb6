VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0B8EB583-E121-11D0-8C51-00C04FC29CEC}#2.0#0"; "Editor.ocx"
Object = "{35DFF50E-DEB3-11D0-8C50-00C04FC29CEC}#1.1#0"; "DropStack.ocx"
Begin VB.Form FEdwina 
   Caption         =   "Edwina"
   ClientHeight    =   6492
   ClientLeft      =   1020
   ClientTop       =   2436
   ClientWidth     =   8400
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   7.2
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Edwina.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6492
   ScaleWidth      =   8400
   Begin MSComctlLib.Toolbar barEdit 
      Align           =   1  'Align Top
      Height          =   312
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   550
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      HelpContextID   =   1241660
      ImageList       =   "imlstEdit"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "Create new file"
            Object.ToolTipText     =   "Create new file"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open existing file"
            Object.ToolTipText     =   "Open existing file"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save file"
            Object.ToolTipText     =   "Save file"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1200
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print file"
            Object.ToolTipText     =   "Print file"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Description     =   "Find text"
            Object.ToolTipText     =   "Find text"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut selection"
            Object.ToolTipText     =   "Cut selection"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy selection"
            Object.ToolTipText     =   "Copy selection"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste from clipboard"
            Object.ToolTipText     =   "Paste from Clipboard"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Description     =   "Undo last edit"
            Object.ToolTipText     =   "Undo last edit"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1200
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Description     =   "Left justify"
            Object.ToolTipText     =   "Left justify"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Description     =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   14
            Style           =   2
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Description     =   "Right justify"
            Object.ToolTipText     =   "Right justify"
            ImageIndex      =   15
            Style           =   2
         EndProperty
      EndProperty
      Begin DropStack.XDropStack dropFind 
         Height          =   288
         Left            =   3912
         TabIndex        =   4
         Top             =   12
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   508
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DropStack.XDropStack dropFile 
         Height          =   288
         Left            =   876
         TabIndex        =   2
         Top             =   12
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   508
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Editor.XEditor edit 
      Height          =   1524
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   1812
      _ExtentX        =   3831
      _ExtentY        =   3408
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   0
      MouseIcon       =   "Edwina.frx":0CFA
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.8
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6240
      TabIndex        =   1
      Top             =   5580
      Visible         =   0   'False
      Width           =   972
   End
   Begin MSComctlLib.StatusBar statEdit 
      Align           =   2  'Align Bottom
      Height          =   456
      Left            =   0
      TabIndex        =   0
      Top             =   6036
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   804
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "Line:  3000 / 3000"
            TextSave        =   "Line:  3000 / 3000"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2611
            MinWidth        =   2611
            Text            =   "Column:  100 / 100"
            TextSave        =   "Column:  100 / 100"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2028
            MinWidth        =   2028
            Text            =   "Percent:  100"
            TextSave        =   "Percent:  100"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1905
            MinWidth        =   1905
            Text            =   "Margin:   90"
            TextSave        =   "Margin:   90"
            Object.ToolTipText     =   "Margin"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "4:54 PM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   970
            MinWidth        =   970
            Text            =   "INS"
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   970
            MinWidth        =   970
            Text            =   "SAV"
            TextSave        =   "SAV"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlstEdit 
      Left            =   5136
      Top             =   408
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":0D16
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":0E28
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":0F3A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":104C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":115E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":1270
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":1382
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":1494
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":15A6
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":16B8
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":17CA
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":18DC
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":1B00
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Edwina.frx":1C12
            Key             =   "Right"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuSepPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuSepExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo                   Ctrl+Z"
      End
      Begin VB.Menu mnuSepCut 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t                        Ctrl+X"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy                   Ctrl+C"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste                  Ctrl+V"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "De&lete                 Del"
      End
      Begin VB.Menu mnuSepSelectAll 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All           Ctrl+A"
      End
      Begin VB.Menu mnuEditTimeDate 
         Caption         =   "Time/&Date"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSearchFindPrevious 
         Caption         =   "Find Pre&vious"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuSearchReplace 
         Caption         =   "&Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuReplaceNext 
         Caption         =   "Re&place Next"
         Shortcut        =   ^{F3}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionFont 
         Caption         =   "Fon&t..."
      End
      Begin VB.Menu mnuOptionSelFont 
         Caption         =   "&Selection Font..."
      End
      Begin VB.Menu mnuOptionBackColor 
         Caption         =   "&Background Color..."
      End
      Begin VB.Menu mnuOptionTextColor 
         Caption         =   "Text &Color..."
      End
      Begin VB.Menu mnuOptionRichText 
         Caption         =   "&Rich Text"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu cmnuContext 
      Caption         =   "Context"
      Begin VB.Menu cmnuCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu cmnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu cmnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu cmnuDelete 
         Caption         =   "De&lete"
      End
      Begin VB.Menu cmnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu cmnuFont 
         Caption         =   "&Font..."
      End
   End
End
Attribute VB_Name = "FEdwina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EPanels
    epLine = 1
    epCol = 2
    epPercent = 3
    epMargin = 4
    epTime = 5
    epIns = 6
    epCap = 7
    epSav = 8
End Enum

Const sBold = "Bold"
Const sItalic = "Italic"
Const sUnderline = "Underline"
Const sLeft = "Left"
Const sCenter = "Center"
Const sRight = "Right"

Private fInput As Boolean
Private fInFindCompleted As Boolean

Private Sub Form_Load()
    cmnuContext.Visible = False
    App.Title = Me.Caption
    If Command$ <> sEmpty Then
        Dim fView As Boolean, fPrint As Boolean, sTarget As String
        ParseCmd fView, fPrint, sTarget
        edit.LoadFile sTarget
        If fView Or fPrint Then edit.Locked = True
        If fPrint Then
            edit.SelPrint Printer.hDC
            Printer.EndDoc
        End If
    End If
    dropFile.MaxCount = 3
    dropFile.Text = edit.FileName
    edit.EnableTab = True
    edit.FindWhatMax = 3
    dropFind.MaxCount = edit.FindWhatMax
    SetTextMode True
    edit_StatusChange 0, 0, 0, 0, 0, 0, edit.DirtyBit
    ' VB won't allow tab in menu editor, so add here
    mnuEditUndo.Caption = "&Undo" & sTab & "Ctrl+Z"
    mnuEditCut.Caption = "Cu&t" & sTab & "Ctrl+X"
    mnuEditCopy.Caption = "&Copy" & sTab & "Ctrl+C"
    mnuEditPaste.Caption = "&Paste" & sTab & "Ctrl+V"
    mnuEditDelete.Caption = "De&lete" & sTab & "Del"
    mnuEditSelectAll.Caption = "Select &All" & sTab & "Ctr+A"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    fInput = False
    If Not edit.DirtyDialog Then Cancel = True
End Sub

Private Sub Form_Resize()
    Dim iMargin As Integer
    Static fInResize As Boolean
    ' Prevent recursive calls
    If fInResize Then Exit Sub
    fInResize = True
    If Me.WindowState = vbMinimized Then Exit Sub
    ' Resize form if edit window is unreasonably small
    If Width < (statEdit.Height * 14) Then Width = statEdit.Height * 14
    If Height < (statEdit.Height * 7) Then Height = statEdit.Height * 7
    edit.Move 0, barEdit.Height, ScaleWidth, ScaleHeight - barEdit.Height - statEdit.Height
    iMargin = Val(Mid$(statEdit.Panels(epMargin).Text, 10))
    edit.RightMargin = (iMargin / 100) * edit.Width
    fInResize = False
End Sub

Private Sub ParseCmd(fView As Boolean, fPrint As Boolean, sTarget As String)
    Dim sToken As String, s As String
    Const sSep = sSpace & sTab & sCrLf
    ' Error handling could be improved
    sToken = GetToken(Command$, sSep)
    Do While sToken <> sEmpty
        Select Case Left$(sToken, 1)
        Case "-", "/"
            s = UCase$(Mid$(sToken, 2, 1))
            If s = "V" Then fView = True
            If s = "P" Then fPrint = True
            ' Throw away unknown options
        Case Else
            ' Use first non-option argument as target
            If sTarget = sEmpty Then sTarget = sToken
        End Select
        sToken = GetToken(sEmpty, sSep)
    Loop
End Sub

Private Sub mnuFileNew_Click()
    If edit.DirtyDialog Then edit.FileNew
    dropFile.Text = edit.FileName
End Sub

Private Sub mnuFileOpen_Click()
    If edit.DirtyDialog Then edit.FileOpen
    dropFile.Text = edit.FileName
    SetTextMode edit.TextMode
End Sub

Private Sub mnuFileSave_Click()
    edit.FileSave
End Sub

Private Sub mnuFileSaveAs_Click()
    edit.FileSaveAs
    dropFile.Text = edit.FileName
End Sub

Private Sub mnuFilePrint_Click()
    edit.FilePrint
End Sub

Private Sub mnuFilePageSetup_Click()
    edit.FilePageSetup
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuEdit_Click()
    mnuEditUndo.Enabled = edit.CanUndo
End Sub

Private Sub mnuEditUndo_Click()
    edit.EditUndo
End Sub

Private Sub mnuEditCopy_Click()
    edit.EditCopy
End Sub

Private Sub mnuEditCut_Click()
    edit.EditCut
End Sub

Private Sub mnuEditDelete_Click()
    edit.EditDelete
End Sub

Private Sub mnuEditPaste_Click()
    edit.EditPaste
End Sub

Private Sub mnuEditSelectAll_Click()
    edit.EditSelectAll
End Sub

Private Sub mnuEditTimeDate_Click()
    edit.SelText = Now
End Sub

Private Sub mnuSearchFind_Click()
    edit.SearchFind
End Sub

Private Sub mnuSearchFindNext_Click()
    edit.SearchFindNext
End Sub

Private Sub mnuSearchFindPrevious_Click()
    Dim i As Integer
    i = edit.SearchOptionDirection
    edit.SearchOptionDirection = 2
    edit.SearchFindNext
    edit.SearchOptionDirection = i
End Sub

Private Sub mnuSearchReplace_Click()
    edit.SearchReplace
End Sub

Private Sub mnuReplaceNext_Click()
    edit.ReplaceNext dropFind, edit.ReplaceWith
    edit.FindNext dropFind
End Sub

Private Sub mnuOptionFont_Click()
    edit.OptionFont
End Sub

Private Sub mnuOptionSelFont_Click()
    edit.OptionFont True
End Sub

Private Sub mnuOptionBackColor_Click()
With edit
    .BackColor = .OptionColor(.BackColor)
End With
End Sub

Private Sub mnuOptionTextColor_Click()
With edit
    If .TextMode Then
        .TextColor = .OptionColor(.TextColor)
    Else
        Dim clr As Long
        clr = IIf(Not IsNull(.SelColor), .SelColor, vbBlack)
        .SelColor = .OptionColor(clr)
    End If
End With
End Sub

Private Sub mnuOptionRichText_Click()
    SetTextMode Not edit.TextMode
End Sub

Private Sub SetTextMode(fText As Boolean)
    edit.TextMode = fText
    mnuOptionRichText.Checked = Not fText
    With barEdit
        .Buttons(sBold).Visible = Not fText
        .Buttons(sItalic).Visible = Not fText
        .Buttons(sUnderline).Visible = Not fText
        .Buttons(sLeft).Visible = Not fText
        .Buttons(sCenter).Visible = Not fText
        .Buttons(sRight).Visible = Not fText
    End With
    mnuOptionSelFont.Enabled = Not fText
    mnuOptionTextColor.Caption = IIf(fText, "Text &Color", "Selection &Color")
End Sub

Private Sub cmnuCopy_Click()
    edit.EditCopy
End Sub

Private Sub cmnuCut_Click()
    edit.EditCut
End Sub

Private Sub cmnuDelete_Click()
    edit.EditDelete
End Sub

Private Sub cmnuPaste_Click()
    edit.EditPaste
End Sub

Private Sub cmnuFont_Click()
    edit.OptionFont
End Sub

Private Sub barEdit_ButtonClick(ByVal Button As Button)
    Select Case Button.Key
    Case "New"
        mnuFileNew_Click
    Case "Open"
        mnuFileOpen_Click
    Case "Save"
        mnuFileSave_Click
    Case "Print"
        edit.SelPrint Printer.hDC
    Case "Find"
        mnuSearchFindNext_Click
    Case "Cut"
        mnuEditCut_Click
    Case "Copy"
        mnuEditCopy_Click
    Case "Paste"
        mnuEditPaste_Click
    Case "Undo"
        mnuEditUndo_Click
    Case sBold
        edit.SelBold = Not edit.SelBold
    Case sItalic
        edit.SelItalic = Not edit.SelItalic
    Case sUnderline
        edit.SelUnderline = Not edit.SelUnderline
    Case sLeft
        edit.SelAlignment = rtfLeft
    Case sCenter
        edit.SelAlignment = rtfCenter
    Case sRight
        edit.SelAlignment = rtfRight
    End Select
End Sub

Private Sub dropFile_Completed(Text As String)
    Static fInCompleted As Boolean
    If fInCompleted Then Exit Sub
    fInCompleted = True
    On Error GoTo FailCompleted
    edit.TextMode = Not IsRTF(Text)
    edit.LoadFile Text
    SetTextMode edit.TextMode
    fInCompleted = False
    Exit Sub
FailCompleted:
    StatusMessage Err.Description & ": " & dropFile.Text
    dropFile.Text = edit.FileName
    fInCompleted = False
End Sub

Private Sub dropFind_Completed(Text As String)
    If fInFindCompleted Then Exit Sub
    fInFindCompleted = True
    edit.FindWhat = Text
    edit.SetFocus
    edit.FindNext
    fInFindCompleted = False
End Sub

Private Sub edit_StatusChange(LineCur As Long, LineCount As Long, _
                              ColumnCur As Long, ColumnCount As Long, _
                              CharacterCur As Long, _
                              CharacterCount As Long, DirtyBit As Boolean)
With statEdit
    .Panels(epLine).Text = "Line:  " & FmtInt(LineCur, 4) & _
                      " / " & FmtInt(LineCount, 4, True)
    .Panels(epCol).Text = "Column:  " & FmtInt(ColumnCur, 3) & _
                      " / " & FmtInt(ColumnCount, 3, True)
    Dim iPercent As Integer
    iPercent = (CharacterCur / (CharacterCount + 1)) * 100
    .Panels(epPercent).Text = "Percent:  " & FmtInt(iPercent, 3)
    .Panels(epSav).Enabled = DirtyBit
    .Panels(epIns).Enabled = Not edit.OverWrite
End With
With barEdit
    .Buttons(sBold).Value = _
        IIf(edit.SelBold, tbrPressed, tbrUnpressed)
    .Buttons(sItalic).Value = _
        IIf(edit.SelItalic, tbrPressed, tbrUnpressed)
    .Buttons(sUnderline).Value = _
        IIf(edit.SelUnderline, tbrPressed, tbrUnpressed)
End With
End Sub

Private Sub edit_SearchChange(Kind As ESearchEvent)
    Static fInSearchChange As Boolean
    If fInSearchChange Then Exit Sub
    fInSearchChange = True
    If Kind = eseFindWhat Then
        If dropFind.Text <> edit.FindWhat Then
            fInFindCompleted = True
            dropFind.Text = edit.FindWhat
            fInFindCompleted = False
        End If
    End If
    fInSearchChange = False
End Sub

Private Sub edit_KeyPress(KeyAscii As Integer)
    statEdit.Style = sbrNormal
End Sub

Private Sub edit_Click()
    statEdit.Style = sbrNormal
End Sub

Private Sub edit_GotFocus()
    statEdit.Style = sbrNormal
End Sub

Private Sub statEdit_PanelClick(ByVal Panel As Panel)
With Panel
    Dim s As String, iVal As Long
    Select Case .Index
    Case epLine
        s = GetPanelInput(Panel, 6, 8)
        iVal = Val(s)
        If iVal >= 1 And iVal <= edit.Lines Then edit.Line = iVal
    Case epCol
        s = GetPanelInput(Panel, 8, 6)
        iVal = Val(s)
        If iVal >= 1 And iVal <= edit.Columns Then edit.Column = iVal
    Case epPercent
        s = GetPanelInput(Panel, 8, 4)
        iVal = Val(s)
        If iVal >= 0 And iVal <= 100 Then edit.Percent = iVal
    Case epMargin
        s = GetPanelInput(Panel, 8, 3)
        iVal = Val(s)
        If iVal <= 0 Then iVal = 90
        edit.RightMargin = (iVal / 100) * edit.Width
        .Text = Mid$(.Text, 1, 10) & iVal
    Case epTime
        .Style = IIf(.Style = sbrTime, sbrDate, sbrTime)
    Case epIns
        edit.OverWrite = Not edit.OverWrite
        statEdit.Panels(epIns).Enabled = Not edit.OverWrite
    Case epCap
        Keyboard.CapsState = Not Keyboard.CapsState
    Case epSav
        edit.DirtyBit = Not edit.DirtyBit
        .Enabled = edit.DirtyBit
    End Select
    On Error Resume Next
    statEdit.Refresh
    edit.SetFocus
End With
End Sub

Function GetPanelInput(pan As Panel, ByVal iField As Long, ByVal cField As Long) As String
With txtInput
    Dim dx As Single, dy As Single, dxIn As Single, dxStart As Single
    dxStart = statEdit.Panels(1).Left
    dx = GetTextExtentWnd(statEdit.hWnd, Left$(pan.Text, iField - 2), dy)
    dxIn = GetTextExtentWnd(statEdit.hWnd, String$(cField, "0"))
    fInput = True
    Set .Font = statEdit.Font
    Dim iPos As Long
    iPos = InStr(iField, pan.Text, "/")
    If iPos Then
        .Text = Trim$(Mid$(pan.Text, iField, iPos - iField - 1))
    Else
        .Text = Trim$(Mid$(pan.Text, iField))
    End If
    .SelStart = 0
    .SelLength = Len(.Text)
    .Left = dxStart + pan.Left + dx * Screen.TwipsPerPixelX
    .Width = dxIn * Screen.TwipsPerPixelX
    .Height = dy * 1.2
    .Top = statEdit.Top + (statEdit.Height / 2) - (.Height / 2)
    .Visible = True
    .SetFocus
    Do While fInput
        DoEvents
    Loop
    GetPanelInput = .Text
    .Visible = False
End With
End Function

Private Sub txtInput_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn, vbKeyEscape, vbKeyTab
        edit.SetFocus
    End Select
End Sub

Private Sub txtInput_LostFocus()
    fInput = False
End Sub

Sub StatusMessage(sMsg As String)
    statEdit.Style = sbrSimple
    statEdit.SimpleText = sMsg
End Sub
