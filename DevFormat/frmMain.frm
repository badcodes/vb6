VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DevStudio 6 Format"
   ClientHeight    =   1020
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   5952
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   5952
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font"
      Height          =   288
      Left            =   4392
      TabIndex        =   2
      Top             =   132
      Width           =   1428
   End
   Begin VB.ComboBox cmbSelect 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   132
      Width           =   4080
   End
   Begin VB.Label lblFont 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Set the Font of MsDevStudio 6.0 Format Option."
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   108
      TabIndex        =   0
      Top             =   576
      Width           =   5724
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSelect_click()
Dim sFontName As String
Dim iFontSize As Long

sFontName = MRegistry.ReadRegKey( _
                        HKEY_CURRENT_USER, _
                        "Software\Microsoft\DevStudio\6.0\Format\" & cmbSelect.Text, _
                        "FontFace", lblFont.Font.Name)
iFontSize = CLng(MRegistry.ReadRegKey( _
                        HKEY_CURRENT_USER, _
                        "Software\Microsoft\DevStudio\6.0\Format\" & cmbSelect.Text, _
                        "FontSize", CStr(lblFont.Font.Size)))
lblFont.Font.Name = sFontName
lblFont.Font.Size = iFontSize

End Sub

Private Sub cmdFont_Click()
Dim CDLG As New CCommonDialogLite
Dim CurFont As New StdFont
Dim fResult As Boolean
CurFont.Name = lblFont.Font.Name
CurFont.Size = lblFont.Font.Size
fResult = CDLG.VBChooseFont(CurFont)
If fResult = False Then Exit Sub

lblFont.Font.Name = CurFont.Name
lblFont.Font.Size = CurFont.Size
lblFont.FontBold = False
lblFont.FontItalic = False
lblFont.FontUnderline = False
lblFont.FontStrikethru = False

If cmbSelect.ListIndex = 0 Then
    Dim i As Integer
    i = cmbSelect.ListCount
    Do While i > 0
    i = i - 1
    MRegistry.WriteRegKey REG_SZ, HKEY_CURRENT_USER, _
                            "Software\Microsoft\DevStudio\6.0\Format\" & cmbSelect.List(i), _
                            "FontFace", CurFont.Name
    MRegistry.WriteRegKey REG_DWORD, HKEY_CURRENT_USER, _
                            "Software\Microsoft\DevStudio\6.0\Format\" & cmbSelect.List(i), _
                            "FontSize", CurFont.Size
    Loop
Else
    MRegistry.WriteRegKey REG_SZ, HKEY_CURRENT_USER, _
                            "Software\Microsoft\DevStudio\6.0\Format\" & cmbSelect.Text, _
                            "FontFace", CurFont.Name
    MRegistry.WriteRegKey REG_DWORD, HKEY_CURRENT_USER, _
                            "Software\Microsoft\DevStudio\6.0\Format\" & cmbSelect.Text, _
                            "FontSize", CurFont.Size
End If



End Sub

Private Sub Form_Load()
With cmbSelect
.AddItem "All Window"
.AddItem "Calls Window"
.AddItem "Disassembly Window"
.AddItem "Memory Window"
.AddItem "Output Window"
.AddItem "Registers Window"
.AddItem "Source Browser"
.AddItem "Source Window"
.AddItem "Variables Window"
.AddItem "Watch Window"
.AddItem "Workspace Window"
End With
cmbSelect.ListIndex = 0
End Sub
