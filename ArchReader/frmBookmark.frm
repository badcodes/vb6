VERSION 5.00
Begin VB.Form frmBookmark 
   Caption         =   "Manage Bookmark"
   ClientHeight    =   2604
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   6732
   LinkTopic       =   "Form1"
   ScaleHeight     =   2604
   ScaleWidth      =   6732
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   330
      Left            =   5385
      TabIndex        =   9
      Top             =   1965
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   330
      Left            =   3900
      TabIndex        =   8
      Top             =   1965
      Width           =   1065
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   330
      Left            =   2424
      TabIndex        =   7
      Top             =   1956
      Width           =   1065
   End
   Begin VB.ComboBox cboIndex 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   315
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2010
      Width           =   855
   End
   Begin VB.TextBox txtZhSubFile 
      Height          =   300
      Left            =   1530
      TabIndex        =   5
      Top             =   1365
      Width           =   4935
   End
   Begin VB.TextBox txtZhFile 
      Height          =   300
      Left            =   1530
      TabIndex        =   3
      Top             =   780
      Width           =   4950
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   225
      Width           =   4935
   End
   Begin VB.Label lblZhSubFile 
      Caption         =   "zhSubFile:"
      Height          =   360
      Left            =   360
      TabIndex        =   4
      Top             =   1425
      Width           =   1215
   End
   Begin VB.Label lblZhFile 
      Caption         =   "zhFile:"
      Height          =   360
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblBMname 
      Caption         =   "Name:"
      Height          =   360
      Left            =   390
      TabIndex        =   0
      Top             =   285
      Width           =   1215
   End
End
Attribute VB_Name = "frmBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim zhBMCollection As typeZhBookmarkCollection

Private Sub cboIndex_Click()

    Dim i As Integer
    Dim pos As Integer
    i = cboIndex.listIndex + 1
    txtName.text = ""
    txtZhFile.text = ""
    txtZhSubFile.text = ""

    With MainFrm
        txtName.text = .mnuBookmark(i).Caption
        pos = InStr(.mnuBookmark(i).Tag, "|")

        If pos > 0 Then
            txtZhFile.text = Left$(.mnuBookmark(i).Tag, pos - 1)
            txtZhSubFile.text = Right$(.mnuBookmark(i).Tag, Len(.mnuBookmark(i).Tag) - pos)
        End If

    End With

End Sub

Private Sub cmdDelete_Click()

    Dim mnuIndex As Integer
    Dim i As Integer

    If cboIndex.ListCount < 1 Then Exit Sub

    If cboIndex.listIndex < 0 Then Exit Sub
    mnuIndex = cboIndex.listIndex + 1
    cboIndex.RemoveItem cboIndex.listIndex

    With MainFrm
        Dim lEnd As Long
        lEnd = .mnuBookmark.Count - 2

        For i = mnuIndex To lEnd
            .mnuBookmark(i).Caption = .mnuBookmark(i + 1).Caption
            .mnuBookmark(i).Tag = .mnuBookmark(i + 1).Tag
        Next

        Unload MainFrm.mnuBookmark(.mnuBookmark.Count - 1)
    End With

    For i = 0 To cboIndex.ListCount - 1
        cboIndex.List(i) = Str$(i + 1)
    Next

    If cboIndex.ListCount > 0 Then cboIndex.listIndex = cboIndex.ListCount - 1 Else Unload Me

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdSave_Click()

    Dim mnuIndex As Integer

    If cboIndex.ListCount < 1 Then Exit Sub

    If cboIndex.listIndex < 0 Then Exit Sub
    mnuIndex = cboIndex.listIndex + 1

    With MainFrm
        .mnuBookmark(mnuIndex).Caption = txtName.text
        .mnuBookmark(mnuIndex).Tag = txtZhFile.text & "|" & txtZhSubFile.text
    End With

    MsgBox "Done!"

End Sub

Private Sub Form_Load()

    Dim zhLocalize As New CLocalize
    zhLocalize.Install Me, LanguageIni
    zhLocalize.loadFormStr
    Set zhLocalize = Nothing
    Dim i As Integer
    Dim pos As Integer

    With MainFrm
        Me.Icon = .Icon
        zhBMCollection.Count = .mnuBookmark.Count - 1
        ReDim zhBMCollection.zhBookmark(.mnuBookmark.Count - 1) As typeZhBookmark
        Dim lEnd As Long
        lEnd = .mnuBookmark.Count - 1

        For i = 1 To lEnd
            cboIndex.AddItem Str$(i)
            zhBMCollection.zhBookmark(i - 1).sName = .mnuBookmark(i).Caption
            pos = InStr(.mnuBookmark(i).Tag, "|")

            If pos > 0 Then
                zhBMCollection.zhBookmark(i - 1).sZhfile = Left$(.mnuBookmark(i).Tag, pos - 1)
                zhBMCollection.zhBookmark(i - 1).sZhsubfile = Right$(.mnuBookmark(i).Tag _
                   , Len(.mnuBookmark(i).Tag) - pos)
            End If

        Next

    End With

    If cboIndex.ListCount > 0 Then cboIndex.listIndex = cboIndex.ListCount - 1

End Sub

