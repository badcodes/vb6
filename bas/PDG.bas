Attribute VB_Name = "PDGAsist"
Option Explicit
Public Type pdginfo
    vailable As Boolean
    isreadonly As Boolean
    name As String '%t
    author As String '%a
    totalpage As String '%p
    download As String '%u
    publisher As String '%c
    pdate As String '%d
    ssid As String '%s
End Type

Public Type PDG

    iszip As Boolean
    isreadonly As Boolean
    infolder As String
    unzipfolder As String
    zipfile As String
    infofile As String
    info As pdginfo

End Type



Public Function GETpdginfo(infofile As String) As pdginfo

    Dim thisbook As pdginfo
    Dim fso As New FileSystemObject
    Dim ff As File
    Dim ft As TextStream
    Dim Num As Integer
    
    Dim tmpstr As String
    Dim str1 As String, str2 As String
    
'    If Dir(bddir(fso.GetParentFolderName(infofile)) + "*.pdg") <> "" Then
'    thisbook.totalpage = 1
'    Do Until Dir() = ""
'    thisbook.totalpage = thisbook.totalpage + 1
'    Loop
'    End If
    
    If Not fso.FileExists(infofile) Then Exit Function
'
'
'    Set ft = fso.CreateTextFile(infofile, True)
'
'    With thisbook
'    .vailable = True
'    .name = InputBox("书名：", "PDGinfo")
'    .author = InputBox("作者：", "PDGinfo")
'    .download = InputBox("下载位置：", "PDGinfo")
'    ft.WriteLine "书名=" + .name
'    ft.WriteLine "作者=" + .author
'    ft.WriteLine "下载位置=" + .download
'    ft.Close
'    End With
'
'    GETpdginfo = thisbook
'    Exit Function
'
'
'    End If
    
        thisbook.vailable = True
  
        Set ff = fso.GetFile(infofile)
        Set ft = ff.OpenAsTextStream(ForReading)
    
        Do Until ft.AtEndOfStream
        tmpstr = ft.ReadLine
        Num = InStr(1, tmpstr, "=")
        If Num > 0 Then
        str1 = RTrim(LTrim(Left(tmpstr, Num - 1)))
        str2 = RTrim(LTrim(Right(tmpstr, Len(tmpstr) - Num)))
        
        Select Case str1
        Case "作者"
            thisbook.author = str2
        Case "书名"
            thisbook.name = str2
        Case "下载位置"
            thisbook.download = str2
        Case "页数"
            thisbook.totalpage = str2
        Case "出版社"
            thisbook.publisher = str2
        Case "出版日期"
            thisbook.pdate = str2
        Case "SS号"
            thisbook.ssid = str2
        End Select
            
        End If
    
        Loop
        ft.Close
        If (GetAttr(infofile) Mod 2) = 1 Then thisbook.isreadonly = True Else thisbook.isreadonly = False
'        If thisbook.isreadonly = False Then
'        If thisbook.author = "" Or thisbook.author = "BEXP" Then
'            thisbook.author = InputBox("No author information for file: " + thisbook.name + Chr(13) + Chr(10) + "Enter it below:", "PDGinfo")
'            Set ft = ff.OpenAsTextStream(ForWriting)
'            ft.WriteLine "书名=" + thisbook.name
'            ft.WriteLine "作者=" + thisbook.author
'            ft.WriteLine "下载位置=" + thisbook.download
'            ft.Close
'
'        End If
'        End If
           GETpdginfo = thisbook

End Function

Public Function bddir(dirname As String) As String
bddir = dirname
If Right(bddir, 1) <> "\" Then bddir = bddir + "\"
End Function

Public Function getpdg(strcatch As String) As PDG

Dim thispdg As PDG
    With thispdg
        .infofile = ""
        .infolder = ""
        .iszip = False
        .unzipfolder = ""
        .zipfile = ""
        .isreadonly = False
    End With
    
    With thispdg.info
        .author = ""
        .download = "'"
        .name = ""
        .totalpage = 0
        .vailable = False
        .isreadonly = False
    End With

    Dim strtype As VbFileAttribute
    Dim fso As New FileSystemObject

    strtype = GetAttr(strcatch)
    
    If strtype = vbDirectory Or strtype = vbDirectory + vbReadOnly Then
        
        If strtype = vbDirectory + vbReadOnly Then thispdg.isreadonly = True
                
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
    If strtype = 32 Or strtype = vbArchive + vbReadOnly Or strtype = vbReadOnly Then
    If strtype = vbArchive + vbReadOnly Then thispdg.isreadonly = True
    If strtype = vbReadOnly Then thispdg.isreadonly = True
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
    

    
    
    getpdg = thispdg
 
End Function



Public Sub checkpdg(thispdg As PDG)
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

Public Function pdgformat(thispdg As pdginfo, formatstr As String) As String
Dim tmpstr As String
tmpstr = formatstr
If MyInstr(formatstr, "%t,%a,%p,%c,%d") = False Then Exit Function

'If InStr(tmpstr, "%title") = 0 And InStr(tmpstr, "%author") = 0 And InStr(tmpstr, "%pages") = 0 Then Exit Function

tmpstr = Replace(tmpstr, "%t", thispdg.name)
tmpstr = Replace(tmpstr, "%a", thispdg.author)
tmpstr = Replace(tmpstr, "%p", thispdg.totalpage)
tmpstr = Replace(tmpstr, "%c", thispdg.publisher)
tmpstr = Replace(tmpstr, "%d", thispdg.pdate)
tmpstr = Replace(tmpstr, "%s", thispdg.ssid)

Dim vStr() As String
Dim i As Long
Dim iL As Long
Dim iU As Long
Dim sPart As String
vStr = Split(tmpstr, "-")
iL = LBound(vStr)
iU = UBound(vStr)

pdgformat = LTrim$(RTrim$(vStr(iL)))
iL = iL + 1
For i = iL To iU
    sPart = LTrim$(RTrim$(vStr(i)))
    If sPart <> "" Then pdgformat = pdgformat & " - " & sPart
Next

pdgformat = Replace$(pdgformat, "()", "")
pdgformat = Replace$(pdgformat, "[]", "")
pdgformat = Replace$(pdgformat, "《》", "")
pdgformat = Replace$(pdgformat, "［］", "")
pdgformat = Replace$(pdgformat, "“”", "")
pdgformat = Replace$(pdgformat, Chr$(34) & Chr$(34), "")
'pdgformat = Replace$(pdgformat, "()", "")


'pdgformat = tmpstr

End Function

