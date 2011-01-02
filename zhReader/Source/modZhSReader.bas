Attribute VB_Name = "MZhReaderViaHttpServer"
Option Explicit

Public Type HttpServerSet 'SectionName "[HttpServer]"
    sName As String
    sVersion As String
    sIP As String
    sPort As String
End Type

Public Type zipUrl
sZipName As String
sHtmlPath As String
sMapPoint As String
End Type

Public zipProtocolHead As String
Public Const zipSep = "|/"
Public Const zipfakeTrail = "[LXRFakeItHoHo]/"
Public Const zipTempName = "zhReader"
Public sTempZip  As String
Public Sub OptionChanged(ByRef formObject As frmOptions)
End Sub

Public Function zipProtocol_ParseURL(ByVal URL As String) As zipUrl
    Dim lPos As Long
    On Error GoTo InvalidURL
    
    URL = linvblib.DecodeUrl(URL, CP_UTF8)
    If StrComp(Left$(URL, Len(zipProtocolHead)), zipProtocolHead, vbTextCompare) = 0 Then
     URL = Right$(URL, Len(URL) - Len(zipProtocolHead))
     End If
    ' Remove the / at the end of the URL

    If Right$(URL, 1) = "/" Then URL = Left$(URL, Len(URL) - 1)
    ' Remove the // at the begining of the URL

    Do While Left$(URL, 1) = "/"
        URL = Mid$(URL, 2)
    Loop

    ' Find the first / from the right
    lPos = InStr(URL, zipSep)

    If lPos > 0 Then

        With zipProtocol_ParseURL
            .sZipName = Left$(URL, lPos - 1)
            .sHtmlPath = Right$(URL, Len(URL) - lPos - Len(zipSep) + 1)
            lPos = InStr(.sHtmlPath, "#")
            If lPos > 0 Then
                .sMapPoint = Right$(.sHtmlPath, Len(.sHtmlPath) - lPos + 1)
                .sHtmlPath = Left$(.sHtmlPath, lPos - 1)
            End If
            
        End With

    End If

InvalidURL:
    Err.Clear
End Function

Public Function isFakeOne(ByRef fakeStr As String, Optional ByRef realStr As String = "") As Boolean
    isFakeOne = False
    realStr = fakeStr
    If Left$(fakeStr, Len(zipfakeTrail)) = zipfakeTrail Then
        isFakeOne = True
        realStr = Right$(fakeStr, Len(fakeStr) - Len(zipfakeTrail))
    End If
End Function




Public Sub getServerSetting(sIniFilename As String, hssSaveTo As HttpServerSet)

With hssSaveTo
    .sName = iniGetSetting(sIniFilename, "HttpServer", "Name")
    .sVersion = iniGetSetting(sIniFilename, "HttpServer", "Version")
    .sIP = iniGetSetting(sIniFilename, "HttpServer", "IP")
    .sPort = iniGetSetting(sIniFilename, "HttpServer", "Port")
End With
    

End Sub
Public Sub saveServerSetting(sIniFilename As String, hssToSave As HttpServerSet)

With hssToSave
    iniSaveSetting sIniFilename, "HttpServer", "Name", .sName
    iniSaveSetting sIniFilename, "HttpServer", "Version", .sVersion
    iniSaveSetting sIniFilename, "HttpServer", "IP", .sIP
    iniSaveSetting sIniFilename, "HttpServer", "Port", .sPort
End With

End Sub

'FIXIT: Declare 'URL' and 'IEView' with an early-bound data type                           FixIT90210ae-R1672-R1B8ZE
Sub eBeforeNavigate(ByRef URL As Variant, ByRef Cancel As Boolean, ByRef IEView As Object, Optional targetFrame As String = "")

     Dim zhUrl As zipUrl
    Dim rZipname As String
    Dim rZhtmname As String
    Dim preTestFile As String
    Dim thefile As String
    
    zhUrl = zipProtocol_ParseURL(URL)

    If zhUrl.sZipName <> "" And zhUrl.sHtmlPath <> "" Then
        
        zhUrl.sHtmlPath = toUnixPath(zhUrl.sHtmlPath)
        
        If isFakeOne(zhUrl.sHtmlPath, rZhtmname) = False Then
            zhrStatus.sCur_zhSubFile = zhUrl.sHtmlPath
        Else
           preTestFile = BuildPath(sTempZip, rZhtmname)
           If linvblib.PathExists(preTestFile) = False Then
                Cancel = True
                MGetView rZhtmname, IEView, targetFrame
                Exit Sub
           End If
           zhrStatus.sCur_zhSubFile = rZhtmname
        End If
        
        zhrStatus.sCur_zhSubFile = toUnixPath(zhrStatus.sCur_zhSubFile)
        
        thefile = zhrStatus.sCur_zhSubFile
        zhrStatus.sCur_zhSubFile = zhrStatus.sCur_zhSubFile & zhUrl.sMapPoint
        If chkFileType(thefile) <> ftIE Then
            Cancel = True
            MGetView thefile, IEView, targetFrame
            Exit Sub
        End If
        
        If Right$(zhrStatus.sCur_zhSubFile, Len(TempHtm)) = TempHtm Then
            zhrStatus.sCur_zhSubFile = Left$(zhrStatus.sCur_zhSubFile, Len(zhrStatus.sCur_zhSubFile) - Len(TempHtm))
        End If
        
        rZipname = LCase$(zhUrl.sZipName)
        rZhtmname = LCase$(zhrStatus.sCur_zhFile)
        If rZipname <> rZhtmname Then
            Cancel = True
            MainFrm.loadzh zhUrl.sZipName, zhUrl.sHtmlPath, True
            Exit Sub
        End If
        

    
        Exit Sub
    End If
  
'    Dim fso As New GFileSystem
'    Dim sBaseDir As String
'    Dim sLocalUrl As String
'    Dim thefile As String
'
'    If fso.PathExists(zhrStatus.sCur_zhFile) = False Then Exit Sub
'
'    sLocalUrl = toUnixPath(CStr(URL))
'    sBaseDir = toUnixPath(sTempZH)
'
'    If InStr(1, LCase$(sLocalUrl), LCase$(sBaseDir), vbTextCompare) <> 1 Then Exit Sub
'
'    If Left$(sLocalUrl, 5) = "file:" And Len(sLocalUrl) > 7 Then sLocalUrl = Right$(sLocalUrl, Len(sLocalUrl) - 8)
'    thefile = Right$(sLocalUrl, Len(sLocalUrl) - Len(sBaseDir) - 1)
'
'    If fso.PathExists(sLocalUrl) = False Then
'        MainFrm.myXUnzip zhrStatus.sCur_zhFile, thefile, sTempZH, zhrStatus.sPWD
'    End If
'
'    If fso.PathExists(sLocalUrl) = False Then Exit Sub
'
'    If chkFileType(thefile) <> ftIE Then
'        Cancel = True
'        MGetView thefile, IEView, targetFrame
'        Exit Sub
'    End If
'
'
'    zhrStatus.sCur_zhSubFile = thefile
'
'    If Right$(thefile, Len(TempHtm)) = TempHtm Then
'        zhrStatus.sCur_zhSubFile = Replace(thefile, TempHtm, "")
'        Exit Sub
'    End If
    

End Sub

'FIXIT: Declare 'URL' and 'IEView' with an early-bound data type                           FixIT90210ae-R1672-R1B8ZE
Sub eNavigateComplete(ByRef URL As Variant, ByRef IEView As Object)

    MainFrm.LeftFrame.Enabled = True
    
    Dim sUrl As String
sUrl = linvblib.UnescapeUrl(CStr(URL))

If zipProtocol_ParseURL(sUrl).sZipName <> "" Then
        MainFrm.AddUniqueItem MainFrm.cmbAddress, zhrStatus.sCur_zhFile
        MainFrm.cmbAddress.text = zhrStatus.sCur_zhFile & zipSep & zhrStatus.sCur_zhSubFile
Else
    MainFrm.cmbAddress.text = sUrl
End If
    
End Sub
'FIXIT: Declare 'IEView' with an early-bound data type                                     FixIT90210ae-R1672-R1B8ZE
Sub eStatusTextChange(ByVal text As String, ByRef IEView As Object)

text = linvblib.UnescapeUrl(text)
MainFrm.StsBar.Panels("ie").text = Replace(text, zipProtocolHead, "")


End Sub

'FIXIT: Declare 'IEView' with an early-bound data type                                     FixIT90210ae-R1672-R1B8ZE
Public Sub MGetView(shortfile As String, ByRef IEView As Object, Optional targetFrame As String = "")

    If shortfile = "" Then MainFrm.appHtmlAbout: Exit Sub
    Dim fso As New GFileSystem
    Dim tempfile As String
    Dim tempFile2 As String
    Dim bUseTemplate As Boolean
    Dim sTemplateFile As String
    Dim ftThis As LNFileType
    Dim sUrl As String
    Dim sFakeUrl As String
    Dim sFakeLink As String
    Dim sBasePart As String
    Dim mapPoint As String
    
    shortfile = linvblib.UnescapeUrl(shortfile)
    tempfile = linvblib.RightLeft(shortfile, "#", vbTextCompare, ReturnOriginalStr)
    mapPoint = linvblib.RightRight(shortfile, "#", vbTextCompare, ReturnEmptyStr)
    shortfile = tempfile
    
    sBasePart = toUnixPath(zipProtocolHead & zhrStatus.sCur_zhFile & zipSep)
    sUrl = sBasePart & toUnixPath(shortfile)
    sFakeLink = sBasePart & zipfakeTrail & toUnixPath(shortfile)
    sFakeUrl = sFakeLink & TempHtm
    
    sTemplateFile = MainFrm.IEView.Tag 'iniGetSetting(, "Viewstyle", "TemplateFile")
    bUseTemplate = (Val(MainFrm.Tag) <> 0) '
    
    ftThis = chkFileType(shortfile)
    Select Case ftThis
    Case ftIE
        If mapPoint = "" Then
            IEView.Navigate2 sUrl, , targetFrame
        Else
            IEView.Navigate2 sUrl & "#" & mapPoint, , targetFrame
        End If
        Exit Sub
    Case ftZIP, ftZhtm
        tempfile = fso.BuildPath(sTempZH, shortfile)
        If fso.PathExists(tempfile) = False Then
            MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, sTempZH, zhrStatus.sPWD
        End If
        If fso.PathExists(tempfile) = False Then Exit Sub
        MainFrm.loadzh tempfile
    Case ftIMG
        tempfile = fso.BuildPath(sTempZip, shortfile)
        tempFile2 = tempfile & TempHtm
        If fso.PathExists(tempfile) = False Then
            MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, sTempZip, zhrStatus.sPWD
        End If
        If fso.PathExists(tempfile) = False Then Exit Sub
        Dim imgToLoad As Picture
        Dim imgHeight As Long
        Dim imgWidth As Long
        Dim screenHeight As Long
        Dim screenWidth As Long
        Dim resizeRateY As Double
        Dim resizeRateX As Double
        Dim resizeRate As Double
        Set imgToLoad = LoadPicture(tempfile)
'FIXIT: MainFrm.ScaleX method has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
        imgHeight = MainFrm.ScaleX(imgToLoad.Height, 8, 3)
'FIXIT: MainFrm.ScaleX method has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
        imgWidth = MainFrm.ScaleX(imgToLoad.Width, 8, 3)
        screenHeight = (IEView.Height - 360) \ Screen.TwipsPerPixelY
        screenWidth = (IEView.Width - 360) \ Screen.TwipsPerPixelX
        Set imgToLoad = Nothing
        resizeRate = 1
        resizeRateY = 1
        resizeRateX = 1
        If imgHeight > screenHeight Then resizeRateY = screenHeight / imgHeight
        If imgWidth > screenWidth Then resizeRateX = screenWidth / imgWidth
        resizeRate = resizeRateY
        If resizeRateY > resizeRateX Then resizeRate = resizeRateX
        If resizeRate < 1 Then
        imgHeight = Int(imgHeight * resizeRate)
        imgWidth = Int(imgWidth * resizeRate)
        Else
        imgHeight = 0
        imgWidth = 0
        End If
        If bUseTemplate Then
            If createHtmlFromTemplate(sFakeLink, sTemplateFile, tempFile2, imgHeight, imgWidth) Then
                IEView.Navigate2 sFakeUrl, , targetFrame
            ElseIf createDefaultHtml(sFakeLink, tempFile2, imgHeight, imgWidth) Then
                IEView.Navigate2 sFakeUrl, , targetFrame
            Else
                IEView.Navigate2 sUrl, , targetFrame
            End If
        ElseIf createDefaultHtml(sFakeLink, tempFile2, imgHeight, imgWidth) Then
            IEView.Navigate2 sFakeUrl, , targetFrame
        Else
            IEView.Navigate2 sUrl, , targetFrame
        End If
    Case ftAUDIO, ftVIDEO
        tempFile2 = fso.BuildPath(sTempZip, shortfile) & TempHtm
        If bUseTemplate Then
            If createHtmlFromTemplate(sUrl, sTemplateFile, tempFile2) Then
                IEView.Navigate2 sFakeUrl, , targetFrame
            ElseIf createDefaultHtml(sUrl, tempFile2) Then
                IEView.Navigate2 sFakeUrl, , targetFrame
            Else
                IEView.Navigate2 sUrl, , targetFrame
            End If
        ElseIf createDefaultHtml(sUrl, tempFile2) Then
            IEView.Navigate2 sFakeUrl, , targetFrame
        Else
            IEView.Navigate2 sUrl, , targetFrame
        End If
    Case Else 'ftTxt
        tempfile = fso.BuildPath(sTempZH, shortfile)
        If fso.PathExists(tempfile) = False Then
            MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, sTempZH, zhrStatus.sPWD
        End If
        If fso.PathExists(tempfile) = False Then Exit Sub
        tempFile2 = fso.BuildPath(sTempZip, shortfile) & TempHtm
        If bUseTemplate Then
            If createHtmlFromTemplate(tempfile, sTemplateFile, tempFile2) Then
                IEView.Navigate2 sFakeUrl, , targetFrame
            ElseIf createDefaultHtml(tempfile, tempFile2) Then
                IEView.Navigate2 sFakeUrl, , targetFrame
            Else
                IEView.Navigate2 tempfile, , targetFrame
            End If
        ElseIf createDefaultHtml(tempfile, tempFile2) Then
            IEView.Navigate2 sFakeUrl, , targetFrame
        Else
            IEView.Navigate2 tempfile, , targetFrame
        End If
'    Case Else
'        tempfile = fso.BuildPath(sTempZH, shortfile)
'        If fso.PathExists(tempfile) = False Then
'            MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, sTempZH, zhrStatus.sPWD
'        End If
'        If fso.PathExists(tempfile) = False Then Exit Sub
'        ShellExecute MainFrm.hwnd, "open", tempfile, "", "", 1
    End Select

End Sub

Public Sub startUP()
   
    sTempZip = BuildPath(Environ$("temp"), zipTempName)
    If PathExists(sTempZip) = False Then MkDir sTempZip
    Load frmServer

End Sub

Public Sub endUP()
    Dim fso As New FileSystemObject
    On Error Resume Next
    fso.DeleteFolder sTempZip, True
    Unload frmServer
End Sub

Public Function zhServer_DecodeUrl(ByVal sUrl As String) As String
Const errUtf8 = 761
Dim errChar As String
Dim ecCount As Long
Dim sTmpUrl As String
'FIXIT: Keyword 'ChrW$' not supported in Visual Basic .NET                                 FixIT90210ae-R6614-H1984
errChar = ChrW$(errUtf8)
ecCount = charCountInStr(sUrl, errChar)
sTmpUrl = DecodeUrl(sUrl, CP_UTF8)
If ecCount = charCountInStr(sTmpUrl, errChar) Then
    zhServer_DecodeUrl = sTmpUrl
Else
    zhServer_DecodeUrl = DecodeUrl(sUrl, 0)
End If

End Function
Sub MNavigate(sUrl As String, IE As WebBrowser, Optional frameName As String = "")

Dim sProtocol As String
Dim sMain As String
Dim sSec As String
Dim sExt As String


sProtocol = linvblib.LeftLeft(sUrl, ":")
If Len(sProtocol) = 1 Then

    sMain = linvblib.LeftLeft(sUrl, "|/", vbBinaryCompare, ReturnOriginalStr)
    sSec = linvblib.LeftRight(sUrl, "|/", vbBinaryCompare, ReturnEmptyStr)
    
    If IsZBook(sMain) Then
        MainFrm.loadzh sMain, sSec
    End If
Else
    IE.Navigate sUrl, , frameName
End If
End Sub
