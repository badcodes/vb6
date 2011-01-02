VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NextPage"
   ClientHeight    =   288
   ClientLeft      =   5820
   ClientTop       =   7920
   ClientWidth     =   6324
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleWidth      =   6324
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   12
      TabIndex        =   0
      Top             =   0
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents WC As CDGSwsHTTP
Attribute WC.VB_VarHelpID = -1

Dim ftmp As String
Dim fNum As Integer
Dim sUrl As String
Dim Hdoc As New HTMLDocument
Dim tagLinks As IHTMLElementCollection
Dim tagLink As IHTMLElement
Dim fso As New FileSystemObject
Dim extList() As String
Dim filenameList() As String
Dim fnListUbound As Long
Dim extListUbound As Long
Public iniFile As String


Private Sub Form_Load()



Dim extToDownload As String
Dim filenameToDownload As String

iniFile = BuildPath(App.Path, "nextpage.ini")

ftmp = Command$
If ftmp <> "" And PathExists(ftmp) Then
    fNum = FreeFile
    Open ftmp For Input As #fNum
    If EOF(fNum) = True Then
        Load frmSetting
        frmSetting.Show 1
        Unload Me
        Exit Sub
    End If
Else
    Load frmSetting
    frmSetting.Show 1
    Unload Me
    Exit Sub
End If





Set WC = New CDGSwsHTTP
extToDownload = iniGetSetting(iniFile, "Download", "Ext")
filenameToDownload = iniGetSetting(iniFile, "Download", "Filename")



extList = Split(extToDownload, "|")
extListUbound = UBound(extList)
filenameList = Split(filenameToDownload, "|")
fnListUbound = UBound(filenameList)

geturl

End Sub

Private Sub geturl()

Dim sExt As String
If EOF(fNum) = True Then
    Close fNum
    Kill ftmp
    Unload Me
    Exit Sub
End If
Line Input #fNum, sUrl


sExt = LCase(GetExtensionName(sUrl))
If sExt = "html" Or sExt = "htm" Or sExt = "aspx" Or sExt = "asp" Then
Me.Caption = "Start Downloading " & sUrl
WC.geturl sUrl
Else
geturl
End If
End Sub

Public Sub addUrl(htmlStr As String)


Dim pList()
Dim pCount As Long
Dim extName As String
Dim lHref As String

ReDim pList(0)
pList(0) = sUrl

Hdoc.body.innerHTML = htmlStr
Set tagLinks = Hdoc.All.tags("a")


Dim baseUrl As String
baseUrl = fso.GetParentFolderName(sUrl)

For Each tagLink In tagLinks

    lHref = tagLink.href
    lHref = Replace(lHref, "about:blank", "", , , vbTextCompare)
    
    If chkExtName(lHref) Then
    
        If LCase(Left(lHref, 5)) <> "http:" Then
            lHref = fso.BuildPath(baseUrl, lHref)
            lHref = fso.GetAbsolutePathName(lHref)
            lHref = Replace(lHref, CurDir & "\", "")
            lHref = Replace(lHref, "\", "/")
            lHref = Replace(lHref, "http://", "http:/")
            lHref = Replace(lHref, "http:/", "http://")
        End If
            pCount = pCount + 1
            ReDim Preserve pList((pCount) * 2)
            pList((pCount - 1) * 2 + 1) = lHref
            pList((pCount - 1) * 2 + 2) = tagLink.innerText
    End If
Next

Set tagLinks = Hdoc.All.tags("img")


baseUrl = fso.GetParentFolderName(sUrl)

For Each tagLink In tagLinks

    lHref = tagLink.src
    lHref = Replace(lHref, "about:blank", "", , , vbTextCompare)
    
    If chkExtName(lHref) Then
    
        If LCase(Left(lHref, 5)) <> "http:" Then
            lHref = fso.BuildPath(baseUrl, lHref)
            lHref = fso.GetAbsolutePathName(lHref)
            lHref = Replace(lHref, CurDir & "\", "")
            lHref = Replace(lHref, "\", "/")
            lHref = Replace(lHref, "http://", "http:/")
            lHref = Replace(lHref, "http:/", "http://")
        End If
            pCount = pCount + 1
            ReDim Preserve pList((pCount) * 2)
            pList((pCount - 1) * 2 + 1) = lHref
            pList((pCount - 1) * 2 + 2) = tagLink.innerText
    End If
Next

        Dim jet As New JetCarNetscape
        If pCount > 1 Then
            jet.AddUrlList pList
        ElseIf pCount = 1 Then
            jet.addUrl pList(1), pList(2), sUrl
        Else
            MsgBox "Not Link found." & "Refer to " & sUrl
        End If

End Sub



Private Sub WC_DownloadComplete()
Me.ProgressBar1.Value = 0
Me.Caption = WC.URL & " Download complete!"
addUrl WC.filedata
geturl
End Sub

Private Sub WC_httpError(errmsg As String, Scode As String)
ProgressBar1 = 0
MsgBox errmsg & vbCrLf & WC.ResponseHeaderString, vbExclamation, "Error"
geturl
End Sub

Private Sub WC_ProgressChanged(ByVal bytesreceived As Long)
Me.Caption = "downloading " & bytesreceived & " bytes received of " & WC.FileSize

' update the progressbar
' I didn't put this code in the module for speed, so if you don't need the progressbar
' remove the following
Dim percentcomplete As Long
percentcomplete = 50 ' so progress shows something if no filesize was returned
If WC.FileSize > 0 Then
   percentcomplete = (bytesreceived / WC.FileSize) * 100
End If
Me.ProgressBar1.Value = percentcomplete
End Sub

Private Function chkExtName(ByVal sHref As String) As Boolean

Dim extName As String
Dim FileName As String

Dim l As Long
sHref = LeftLeft(sHref, "?", vbBinaryCompare, ReturnOriginalStr)

extName = LCase(GetExtensionName(sHref))
FileName = LCase(GetBaseName(sHref))

For l = 0 To extListUbound
If extName = extList(l) Then chkExtName = True: Exit Function
Next

For l = 0 To fnListUbound
If FileName = filenameList(l) Then chkExtName = True: Exit Function
Next



End Function
