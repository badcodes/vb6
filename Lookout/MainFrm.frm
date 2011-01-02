VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lookout"
   ClientHeight    =   480
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3960
      Top             =   0
   End
   Begin VB.CommandButton cmdSetting 
      Caption         =   "¡ú"
      Height          =   280
      Left            =   2040
      TabIndex        =   3
      Top             =   90
      Width           =   375
   End
   Begin VB.ComboBox cboHistory 
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Top             =   90
      Width           =   5175
   End
   Begin VB.ComboBox cboEngine 
      Height          =   300
      ItemData        =   "MainFrm.frx":058A
      Left            =   120
      List            =   "MainFrm.frx":058C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton cmdLookout 
      Caption         =   "Lookout"
      Default         =   -1  'True
      Height          =   300
      Left            =   7800
      TabIndex        =   2
      Top             =   75
      Width           =   975
   End
   Begin VB.Menu mnuEngine 
      Caption         =   "Engine"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu eEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu eADD 
         Caption         =   "Add"
      End
      Begin VB.Menu eRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuSp 
         Caption         =   "©¥©¥©¥©¥©¥©¥©¥"
         Enabled         =   0   'False
      End
      Begin VB.Menu Tclear 
         Caption         =   "Clear History"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tSE() As searchEngine
Dim lEngine As Long
Dim lHistory As Long


Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then Combo2.Text = "": Exit Sub
If KeyCode = vbKeyDelete And Shift = 2 Then Tclear_Click: Exit Sub

End Sub

Private Sub Command1_Click()
If Combo2.Text = "" Then Exit Sub
Dim searchurl As String
Dim queryText As String
queryText = Combo2.Text
searchurl = Replace(e(Combo1.ListIndex + 1), "#####query#####", queryText)
Shell Environ("programfiles") + "\Internet Explorer\iexplore.exe " + searchurl, vbNormalFocus
For i = 0 To Combo2.ListCount - 1
If Combo2.List(i) = queryText Then Combo2.RemoveItem i: Exit For
Next
Combo2.AddItem queryText
Combo2.Text = queryText
End Sub






Private Sub Command3_Click()
MainFrm.PopupMenu mnuEngine, , Command3.Left + Command3.Width / 2, Command3.Top
End Sub

Private Sub eADD_Click()
Dim ename As String
Dim elink As String
ename = InputBox("Input the Name of the Search Engine :", App.ProductName, , Command3.Left + Command3.Width / 2, Command3.Top)
If ename = "" Then Exit Sub
elink = InputBox("Input the Link of the Search Engine." + vbCrLf + "Remember Fullfill the SearchContont Width " + Chr(34) + "#####query#####" + Chr(34) + " .", "WebSearch", , Command3.Left + Command3.Width / 2, Command3.Top)
If elink = "" Then Exit Sub
Combo1.AddItem ename
e.Add elink
Combo1.ListIndex = Combo1.ListCount - 1
Esave
End Sub

Private Sub eEdit_Click()
Dim ename As String
Dim elink As String
If Combo1.ListIndex < 0 Then Exit Sub
ei = Combo1.ListIndex
ename = Combo1.List(ei)
elink = e(ei + 1)
ename = InputBox("Input the Name of the Search Engine :", "WebSearch", ename)
If ename = "" Then Exit Sub
elink = InputBox("Input the Link of the Search Engine." + vbCrLf + "Fullfill the SearchContent With " + Chr(34) + "#####query#####" + Chr(34) + " .", "WebSearch", elink)
If elink = "" Then Exit Sub
Combo1.RemoveItem ei
e.Remove ei + 1
Combo1.AddItem ename
e.Add elink
Combo1.ListIndex = Combo1.ListCount - 1
Esave
End Sub

Private Sub eRemove_Click()
If Combo1.ListIndex < 0 Then Exit Sub
ei = Combo1.ListIndex
Combo1.RemoveItem ei
e.Remove ei + 1
Esave
End Sub

Private Sub Form_Load()

Dim fso As New FileSystemObject

MainFrm.Caption = App.ProductName
sLookoutInI = fso.BuildPath(App.Path, "config.ini")
If fso.FileExists(sLookoutInI) = False Then fso.CreateTextFile sLookoutInI, True
loadIni

Dim tempstr As String
Dim tnum As Integer
Dim Num As Integer
Dim defaultEN As Integer


For i = e.Count To 1 Step -1
e.Remove i
Next
Num = -1

For i = 0 To 99
tempstr = GetSetting(App.ProductName, "Engine", LTrim(Str(i)))

If tempstr <> "" Then
    Dim pos As Integer
    pos = InStr(tempstr, "|")
    If pos > 0 Then
    Num = Num + 1
    Combo1.AddItem Left(tempstr, pos - 1)
    e.Add Right(tempstr, Len(tempstr) - pos)
    End If
End If

Next

tnum = Val(GetSetting(App.ProductName, "History", "Count"))

For i = tnum - 1 To 0 Step -1
tempstr = GetSetting(App.ProductName, "History", LTrim(Str(i)))

If tempstr <> "" Then
Combo2.AddItem tempstr
End If
Next

defaultEN = Val(GetSetting(App.ProductName, "Engine", "DefaultEngine"))
If defaultEN < Combo1.ListCount And defaultEN >= 0 Then Combo1.ListIndex = defaultEN
'SetWindowPos Me.hwnd, hwnd_topmost, Me.Left, Me.Top, Me.Width, Me.Height, swp_nosize + swp_nomove

End Sub



Private Sub Form_Resize()

If Me.WindowState = 1 Then Timer1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)


iniSaveSetting sLookoutInI, "History", "Count", Combo2.ListCount

For i = 0 To Combo2.ListCount - 1
iniSaveSetting sLookoutInI, "History", LTrim(Str(i)), Combo2.List(i)

Next
iniSaveSetting sLookoutInI, "Engine", "DefaultEngine", LTrim(Str(Combo1.ListIndex))
End Sub

Sub Esave()
ecount = e.Count
DeleteSetting App.ProductName, "Engine"
For i = 1 To ecount
iniSaveSetting sLookoutInI, "Engine", LTrim(Str(i - 1)), Combo1.List(i - 1) + "|" + e(i)
Next
End Sub



Private Sub Tclear_Click()
If Combo2.ListCount = 0 Then Exit Sub
Combo2.Clear
DeleteSetting App.ProductName, "History"
End Sub

Private Sub Timer1_Timer()
Load trayform
trayform.thetext = Combo2.Text
Unload Me
End Sub



