Attribute VB_Name = "Module1"

Public Type pdginfo
    vailable As Boolean
    name As String
    author As String
    totalpage As Integer
    download As String
End Type

Public Type pdg

    iszip As Boolean
    infolder As String
    unzipfolder As String
    zipfile As String
    infofile As String
    info As pdginfo

End Type
Public thispdg As pdg


Public Function GETpdginfo(infofile As String) As pdginfo

    Dim thisbook As pdginfo
    Dim fso As New FileSystemObject
    Dim ff As File
    Dim ft As TextStream
    Dim num As Integer
    If Not fso.FileExists(infofile) Then Exit Function
    
        thisbook.vailable = True
        Set ff = fso.GetFile(infofile)
        Set ft = ff.OpenAsTextStream(ForReading)
    
        Do Until ft.AtEndOfStream
        tmpstr = ft.ReadLine
        num = InStr(1, tmpstr, "=")
        If num > 0 Then
        str1 = RTrim(LTrim(Left(tmpstr, num - 1)))
        str2 = RTrim(LTrim(Right(tmpstr, Len(tmpstr) - num)))
        
        Select Case str1
        Case "页数"
            thisbook.totalpage = Val(str2)
        Case "作者"
            thisbook.author = str2
        Case "书名"
            thisbook.name = str2
        Case "下载位置"
            thisbook.download = str2
        End Select
            
        End If
    
        Loop
        ft.Close
        If thisbook.author = "" Or thisbook.author = "BEXP" Then
            thisbook.author = InputBox("No author information for file: " + thisbook.name + Chr(13) + Chr(10) + "Enter it below:", "PDGinfo")
            Set ft = ff.OpenAsTextStream(ForWriting)
            ft.WriteLine "书名=" + thisbook.name
            ft.WriteLine "作者=" + thisbook.author
            ft.WriteLine "页数=" + Str(thisbook.totalpage)
            ft.WriteLine "下载位置=" + thisbook.download
            ft.Close
            
        End If
           GETpdginfo = thisbook

End Function

Public Function bddir(dirname As String) As String
bddir = dirname
If Right(bddir, 1) <> "\" Then bddir = bddir + "\"
End Function

Public Sub getpdg(strcatch As String)


    With thispdg
        .infofile = ""
        .infolder = ""
        .iszip = False
        .unzipfolder = ""
        .zipfile = ""
    End With
    
    With thispdg.info
        .author = ""
        .download = "'"
        .name = ""
        .totalpage = 0
        .vailable = False
    End With

    Dim strtype As VbFileAttribute
    Dim fso As New FileSystemObject

    strtype = GetAttr(strcatch)
    
    If strtype = vbDirectory Then
                
        thispdg.infolder = strcatch
        strcatch = bddir(strcatch)
        
        If Dir(strcatch + "*.pdg") <> "" Then
            thispdg.infofile = strcatch + "BOOKINFO.DAT"
        ElseIf Dir(strcatch + "*.ZIP") <> "" Then
            thispdg.unzipfolder = Environ("temp") + "\PdgZF"
            thispdg.zipfile = strcatch + Dir(strcatch + "*.zip")
            thispdg.infofile = strcatch + "bookinfo.dat"
            thispdg.iszip = True
        Else
            MsgBox ("Error:Not pdg folder or zipfile")
            End
        End If
        
    End If
    If strtype = 32 Then
    If LCase(fso.GetExtensionName(strcatch)) <> "zip" Then
    MsgBox ("Error:NOT pdg folder or zipfile")
    End
    End If
    thispdg.infolder = fso.GetParentFolderName(strcatch)
    thispdg.unzipfolder = Environ("temp") + "\PdgZF"
    thispdg.zipfile = strcatch
    thispdg.infofile = bddir(thispdg.infolder) + "bookinfo.dat"
    thispdg.iszip = True
    End If
    thispdg.info = GETpdginfo(thispdg.infofile)
    Dir Environ("temp")
 
End Sub

Public Function createini()
Dim ssreader As String
Dim tmpstr As String
ssreader = InputBox("输入PDG阅读器的地址")
If Dir(ssreader) = "" Then
 MsgBox "Error: 找不到PDG阅读器"
 End
End If

tmpstr = App.Path
If Right(tmpstr, 1) <> "\" Then tmpstr = tmpstr + "\"
Open tmpstr + App.EXEName + ".ini" For Output As #1
Print #1, "path=" + ssreader
Close #1

createini = tmpstr + App.EXEName + ".ini"

End Function

Public Sub checkpdg()
Dim fso As New FileSystemObject
With thispdg

If .infolder = "" Then MsgBox "CHECK PDG ERROR": Exit Sub
If .infofile = "" Then MsgBox "CHECK PDG ERROR": Exit Sub
If Not fso.FolderExists(.infolder) Then MsgBox "CHECK PDG ERROR": Exit Sub
If Not fso.FileExists(.infofile) Then MsgBox "CHECK PDG ERROR": Exit Sub
If .iszip Then
    If .zipfile = "" Then MsgBox "CHECK PDG ERROR": Exit Sub
    If .unzipfolder = "" Then MsgBox "CHECK PDG ERROR": Exit Sub
    If Not fso.FileExists(.zipfile) Then MsgBox "CHECK PDG ERROR": Exit Sub
End If


End With
End Sub
