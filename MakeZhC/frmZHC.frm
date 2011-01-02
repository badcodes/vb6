VERSION 5.00
Begin VB.Form frmZHC 
   Caption         =   "Make ZH Comment"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMake 
      Caption         =   "Make"
      Default         =   -1  'True
      Height          =   495
      Left            =   5025
      TabIndex        =   18
      Top             =   5100
      Width           =   1695
   End
   Begin VB.TextBox txtBaseFolder 
      Height          =   300
      Left            =   1305
      TabIndex        =   16
      Top             =   4515
      Width           =   5415
   End
   Begin VB.TextBox txtContentFile 
      Height          =   300
      Left            =   1305
      TabIndex        =   14
      Top             =   3915
      Width           =   5415
   End
   Begin VB.TextBox txtDefaultFile 
      Height          =   300
      Left            =   1305
      TabIndex        =   12
      Top             =   3315
      Width           =   5415
   End
   Begin VB.CheckBox chkShowMenu 
      Caption         =   "ShowMenu"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CheckBox chkShowLeft 
      Caption         =   "ShowLeft"
      Height          =   375
      Left            =   465
      TabIndex        =   10
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Left            =   1305
      TabIndex        =   9
      Top             =   2715
      Width           =   5415
   End
   Begin VB.TextBox txtCatalog 
      Height          =   300
      Left            =   1305
      TabIndex        =   7
      Top             =   2115
      Width           =   5415
   End
   Begin VB.TextBox txtPublisher 
      Height          =   300
      Left            =   1305
      TabIndex        =   5
      Top             =   1515
      Width           =   5415
   End
   Begin VB.TextBox txtAuthor 
      Height          =   300
      Left            =   1305
      TabIndex        =   3
      Top             =   915
      Width           =   5415
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   1305
      TabIndex        =   1
      Top             =   315
      Width           =   5415
   End
   Begin VB.Label Label8 
      Caption         =   "BaseFolder"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "ContentFile"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "DefaultFile"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Date"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Catalog"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Publisher"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Author"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmZHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject
Dim ZHC As New clsZhComment



Private Sub cmdMake_Click()

Dim sContentFilePath As String
Dim sArrcontent() As String
Dim lContentCount As Long


ZHC.selfReset
With ZHC
.sTitle = txtTitle.Text
.sAuthor = txtAuthor.Text
.sCatalog = txtCatalog.Text
.sDate = txtDate.Text
.sDefaultfile = txtDefaultFile.Text
.sPublisher = txtPublisher.Text
.bShowLeft = chkShowLeft.Value
.bShowMenu = chkShowMenu.Value
End With


sContentFilePath = fso.BuildPath(txtBaseFolder.Text, txtContentFile.Text)
If fso.FileExists(sContentFilePath) = True Then

           getlinks sContentFilePath, sArrcontent(), lContentCount
           sArrcontent(0, 0) = ZHC.sTitle + "\"
           sArrcontent(1, 0) = ZHC.sDefaultfile
           ZHC.CopyContentFrom sArrcontent()
   
End If

Dim zhcfile As String

If txtBaseFolder.Text = "" Then zhcfile = App.Path Else zhcfile = txtBaseFolder.Text
zhcfile = fso.BuildPath(zhcfile, zhCommentFileName)
ZHC.saveZhCommentToFile zhcfile
MsgBox "Done!"
End Sub

Private Sub Form_Load()
If Command$ = "" Then Exit Sub
Dim cmdLine As String
Dim projectname As String
cmdLine = Command$
txtDate.Text = date$

If fso.FolderExists(cmdLine) Then
    txtBaseFolder.Text = cmdLine
    txtTitle.Text = fso.GetBaseName(cmdLine)
ElseIf fso.FileExists(cmdLine) Then
    txtBaseFolder.Text = fso.GetParentFolderName(cmdLine)
    txtTitle.Text = fso.GetBaseName(txtBaseFolder.Text)
    txtDefaultFile.Text = fso.GetFileName(cmdLine)
    txtContentFile.Text = txtDefaultFile.Text
End If

End Sub
Function getlinks(htmlfile As String, links() As String, linknum As Long) As Boolean

Dim HtmDoc As New HTMLDocument
Dim theHtm As IHTMLDocument2

Set theHtm = HtmDoc.createDocumentFromUrl(htmlfile, "")
Do Until theHtm.readyState = "complete"
DoEvents
Loop

linknum = 1
Dim A_C As IHTMLElementCollection
Set A_C = theHtm.All.tags("A")
For i = 0 To A_C.length - 1
tempstr = A_C(i).href

If Left(tempstr, 8) = "file:///" Then
    ReDim Preserve links(1, linknum) As String
    links(0, linknum) = ZHC.sTitle + "\" + A_C(i).innerText
    links(1, linknum) = fso.BuildPath(fso.GetParentFolderName(txtContentFile.Text), gethref(A_C(i).outerHTML))
    If Left(links(1, linknum), 1) = "#" Then links(1, linknum) = fso.GetFileName(htmlfile) + links(1, linknum)
    linknum = linknum + 1
End If

Next


End Function

Function gethref(marka As String) As String

Dim SrcStr As String
Dim tempchar As String
SrcStr = LCase(marka)

n = InStr(SrcStr, "href=" + Chr(34))

If n > 0 Then
    n = n + 6
    tempchar = Mid(marka, n, 1)
    Do Until tempchar = Chr(34) Or n > Len(marka)
    gethref = gethref + tempchar
    n = n + 1
    tempchar = Mid(marka, n, 1)
    Loop
    
End If

'gethref = Chr(34) + gethref + Chr(34)
End Function
