Attribute VB_Name = "MZhReaderViaProtocol"
Option Explicit
'Private zhReaderTemp As New clsTempFile
'Private bReallyComplete As Boolean
Public sTempZip As String

Sub eBeforeNavigate(ByRef URL As Variant, ByRef Cancel As Boolean, ByRef IEView As WebBrowser, Optional ByVal Frame As String)

    Dim zhUrl As zipUrl
    Dim rZipname As String
    Dim rZhtmname As String
    Dim preTestFile As String
    Dim sPureName As String
    Dim bFakeUrl As Boolean
    zhUrl = zipProtocol_ParseURL(URL)

    If zhUrl.sZipName <> "" And zhUrl.sHtmlPath <> "" Then
        zhUrl.sHtmlPath = toUnixPath(zhUrl.sHtmlPath)
        bFakeUrl = FakeToReal(zhUrl.sHtmlPath, zhUrl.sHtmlPath)
        sPureName = zhUrl.sHtmlPath

        If Right$(zhUrl.sHtmlPath, Len(TempHtm)) = TempHtm Then
            sPureName = Left$(zhUrl.sHtmlPath, Len(zhUrl.sHtmlPath) - Len(TempHtm))
        End If

        rZipname = LCase$(zhUrl.sZipName)
        rZhtmname = LCase$(zhrStatus.sCur_zhFile)

        If rZipname <> rZhtmname Then
            Cancel = True
            MainFrm.loadzh zhUrl.sZipName, sPureName, True
            Exit Sub
        End If

        'If bFakeUrl Then
        preTestFile = BuildPath(sTempZip, zhUrl.sHtmlPath)

        If linvblib.PathExists(preTestFile) = False Then
            Cancel = True
            MGetView sPureName, IEView, Frame
            Exit Sub
        End If

        'End If
        zhrStatus.sCur_zhSubFile = sPureName
        Exit Sub
    End If

    Dim fso As New gCFileSystem
    Dim sBaseDir As String
    Dim sLocalUrl As String
    Dim thefile As String

    If fso.PathExists(zhrStatus.sCur_zhFile) = False Then Exit Sub
    sLocalUrl = toUnixPath(CStr(URL))
    sBaseDir = toUnixPath(sTempZH)

    If InStr(1, LCase$(sLocalUrl), LCase$(sBaseDir), vbTextCompare) <> 1 Then Exit Sub

    If Left$(sLocalUrl, 5) = "file:" And Len(sLocalUrl) > 7 Then sLocalUrl = Right$(sLocalUrl, Len(sLocalUrl) - 8)
    thefile = Right$(sLocalUrl, Len(sLocalUrl) - Len(sBaseDir) - 1)

    If fso.PathExists(sLocalUrl) = False Then
        MainFrm.myXUnzip zhrStatus.sCur_zhFile, thefile, sTempZH, zhrStatus.sPWD
    End If

    If fso.PathExists(sLocalUrl) = False Then Exit Sub
    '    If chkFileType(thefile) <> ftIE Then
    '
    '        Cancel = True
    '        mGetView thefile, IEView
    '        Exit Sub
    '
    '    End If
    zhrStatus.sCur_zhSubFile = thefile

    If Right$(thefile, Len(TempHtm)) = TempHtm Then
        zhrStatus.sCur_zhSubFile = Replace(thefile, TempHtm, "")
        Exit Sub
    End If

End Sub

Sub eNavigateComplete(ByRef URL As Variant, ByRef IEView As WebBrowser)

'    Dim sUrl As String
'    sUrl = CStr(URL)
'

'    Else
'        MainFrm.cmbAddress.text = linvblib.UnescapeUrl(IEView.LocationURL)
'    End If
Dim sUrl As String, sTemp As String
Dim sPath As String
'URL = linvblib.UnescapeUrl(CStr(URL))
sUrl = linvblib.UnescapeUrl(CStr(URL))
sPath = linvblib.toUnixPath(sUrl)
sTemp = linvblib.toUnixPath(sTempZH)

If sTemp = "" Then Exit Sub
If Left$(sPath, Len(sTemp)) = sTemp Then
    sPath = Right$(sPath, Len(sPath) - Len(sTemp))
    sPath = zhrStatus.sCur_zhFile & "|" & sPath
    If Right$(sPath, Len(TempHtm)) = TempHtm Then sPath = Left$(sPath, Len(sPath) - Len(TempHtm))
    MainFrm.AddUniqueItem MainFrm.cmbAddress, zhrStatus.sCur_zhFile '.cmbAddress, spath
    MainFrm.cmbAddress.text = sPath
ElseIf MUseZipProtocol.zipProtocol_ParseURL(sUrl).sZipName <> "" Then
        sUrl = toUnixPath(zipProtocolHead & zhrStatus.sCur_zhFile & zipSep)
        sUrl = sUrl & zhrStatus.sCur_zhSubFile ' sBasePart & toUnixPath(shortfile)
        MainFrm.cmbAddress.text = sUrl 'zhrStatus.sCur_zhFile & "/" & zhrStatus.sCur_zhSubFile
Else
    MainFrm.cmbAddress.text = sUrl
End If

End Sub

Sub eStatusTextChange(ByVal text As String, ByRef IEView As WebBrowser)

    MainFrm.StsBar.Panels("ie").text = text

End Sub

Public Sub MGetView(shortfile As String, ByRef IEView As WebBrowser, Optional ByVal Frame As String)

    Dim sRealPath As String

    If shortfile = "" Then MainFrm.appHtmlAbout: Exit Sub

    If Right$(shortfile, 1) = "/" Or Right$(shortfile, 1) = "\" Then
        zhrStatus.sCur_zhSubFile = shortfile
        Exit Sub
    End If

    Dim fso As New gCFileSystem
    Dim tempfile As String
    Dim tempFile2 As String
    Dim bUseTemplate As Boolean
    Dim sTemplateFile As String
    Dim sExt As String
    Dim sUrl As String
    Dim sFakeUrl As String
    Dim sFakeLink As String
    Dim sBasePart As String

    shortfile = linvblib.UnescapeUrl(shortfile)
    sBasePart = toUnixPath(zipProtocolHead & zhrStatus.sCur_zhFile & zipSep)
    sUrl = sBasePart & toUnixPath(shortfile)
    sFakeLink = sBasePart & zipfakeTrail & toUnixPath(shortfile)
    sFakeUrl = sFakeLink & TempHtm
    
    sTemplateFile = MainFrm.IEView.Tag ' hIni.GetSetting("Viewstyle", "TemplateFile")
    bUseTemplate = (Val(MainFrm.Tag) <> 0) 'CBoolStr(hIni.GetSetting("ViewStyle", "UseTemplate"))
    

    zhrStatus.sCur_zhSubFile = shortfile
    sRealPath = RightLeft(shortfile, "#", vbBinaryCompare, ReturnOriginalStr)
    sExt = LCase$(fso.GetExtensionName(sRealPath))
    MainFrm.NotPreOperate = True

    Select Case sExt
    Case "html", "htm", "mhtm", "shtm", "asp", "pdf", "swf", "ico", "gif"
        IEView.Navigate sUrl, , Frame
        Exit Sub
    Case "zip", "zhtm", "zjpg" 'ftZIP, ftZhtm
        tempfile = fso.BuildPath(sTempZH, shortfile)

        If fso.PathExists(tempfile) = False Then
            MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, sTempZH, zhrStatus.sPWD
        End If

        If fso.PathExists(tempfile) = False Then Exit Sub
        MainFrm.loadzh tempfile
    Case "jpg", "jpeg", "jpe", "bmp"
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
        On Error GoTo Herr
        Set imgToLoad = LoadPicture(tempfile)
        imgHeight = MainFrm.ScaleX(imgToLoad.Height, 8, 3)
        imgWidth = MainFrm.ScaleX(imgToLoad.Width, 8, 3)
        screenHeight = (IEView.Height - 360) \ Screen.TwipsPerPixelY
        screenWidth = (IEView.Width - 480) \ Screen.TwipsPerPixelX
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

        'MainFrm.NotPreOperate = False

        If bUseTemplate Then

            If createHtmlFromTemplate(sFakeLink, sTemplateFile, tempFile2, imgHeight, imgWidth) Then
                IEView.Navigate sFakeUrl, , Frame
            ElseIf createDefaultHtml(sFakeLink, tempFile2, imgHeight, imgWidth) Then
                IEView.Navigate sFakeUrl, , Frame
            Else
                IEView.Navigate sUrl, , Frame
            End If

        ElseIf createDefaultHtml(sFakeLink, tempFile2, imgHeight, imgWidth) Then
            IEView.Navigate sFakeUrl, , Frame
        Else
            IEView.Navigate sUrl, , Frame
        End If

    Case "txt", "ini", "vbp", "cfg", "bat", "hhp", "hhc", "txtindex"
        tempfile = fso.BuildPath(sTempZH, shortfile)

        If fso.PathExists(tempfile) = False Then
            MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, sTempZH, zhrStatus.sPWD
        End If

        If fso.PathExists(tempfile) = False Then Exit Sub
        tempFile2 = fso.BuildPath(sTempZip, shortfile) & TempHtm

        If bUseTemplate Then

            If createHtmlFromTemplate(tempfile, sTemplateFile, tempFile2) Then
                IEView.Navigate sFakeUrl, , Frame
            ElseIf createDefaultHtml(tempfile, tempFile2) Then
                IEView.Navigate sFakeUrl, , Frame
            Else
                IEView.Navigate tempfile, , Frame
            End If

        ElseIf createDefaultHtml(tempfile, tempFile2) Then
            IEView.Navigate sFakeUrl, , Frame
        Else
            IEView.Navigate tempfile, , Frame
        End If

    Case "mp3", "mpg", "mpeg", "avi", "wma", "wmv", "wav", "rm", "rmvb"
        tempFile2 = fso.BuildPath(sTempZip, shortfile) & TempHtm

        If bUseTemplate Then

            If createHtmlFromTemplate(sUrl, sTemplateFile, tempFile2) Then
                IEView.Navigate sFakeUrl, , Frame
            ElseIf createDefaultHtml(sUrl, tempFile2) Then
                IEView.Navigate sFakeUrl, , Frame
            Else
                IEView.Navigate sUrl, , Frame
            End If

        ElseIf createDefaultHtml(sUrl, tempFile2) Then
            IEView.Navigate sFakeUrl, , Frame
        Else
            IEView.Navigate sUrl, , Frame
        End If

    Case Else
        tempfile = fso.BuildPath(sTempZH, shortfile)

        If fso.PathExists(tempfile) = False Then
            MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, sTempZH, zhrStatus.sPWD
        End If

        If fso.PathExists(tempfile) = False Then Exit Sub
        ShellExecute MainFrm.hwnd, "open", tempfile, "", "", 1
    End Select

    Exit Sub
Herr:
    MsgBox "Error Number:" & Err.Number & vbCrLf & "Description:" & Err.Description, vbMsgBoxHelpButton, Err.Source, Err.HelpFile, Err.HelpContext

End Sub

Public Sub startUP()

    sTempZip = BuildPath(Environ$("temp"), zipTempName)

    If PathExists(sTempZip) = False Then MkDir sTempZip

End Sub

Public Sub endUP()

    If App.PrevInstance Then Exit Sub
    Dim fso As New FileSystemObject
    On Error Resume Next
    fso.DeleteFolder sTempZip, True

End Sub
'Public Function getTagsProperty(ByRef theHtm As IHTMLDocument2, ByVal tagName As String, ByVal propertyName As String, sBaseUrl As String, ByVal sBaseDir As String) As String
'
'    Dim fso As New FileSystemObject
'    Dim ltagCount As Long
'    Dim lLoop As Long
'    Dim sPropertyValue As String
'    Dim stmp As String
'    Dim iheTag As IHTMLElement
'
'    If theHtm.ReadyState = "uninitialized" Then Exit Function
'
'    ltagCount = theHtm.All.tags(tagName).length
'
'    If ltagCount < 1 Then Exit Function
'
'    Dim lEnd As Long
'    lEnd = ltagCount - 1
'
'    For lLoop = 0 To lEnd
'
'        Set iheTag = theHtm.All.tags(tagName).item(lLoop)
'
'        If IsNull(iheTag.getAttribute(propertyName, 2)) Then GoTo forContune
'        sPropertyValue = iheTag.getAttribute(propertyName, 2)
'
'        If sPropertyValue = "" Then GoTo forContune
'        'If InStr(sPropertyValue, "%") > 0 Then sPropertyValue = DecodeURL(sPropertyValue, CP_UTF8)
'        stmp = Replace(LCase$(sPropertyValue), "\", "/")
'        stmp = Replace(stmp, "%20", " ")
'
'        'If left$(sTmp, 8) = "file:///" Then GoTo forContune
'
'        If InStr(stmp, ":") > 0 Then GoTo forContune
'
'        stmp = fso.BuildPath(fso.GetParentFolderName(sBaseUrl), stmp)
'        stmp = fso.GetAbsolutePathName(stmp)
'        stmp = replaceSlash(stmp)
'        sBaseDir = replaceSlash(sBaseDir)
'
'        If InStr(1, stmp, sBaseDir, vbTextCompare) = 1 Then
'            stmp = Right$(stmp, Len(stmp) - Len(sBaseDir) - 1)
'            getTagsProperty = getTagsProperty & "," & stmp
'        End If
'
'forContune:
'    Next
'
'    If Left$(getTagsProperty, 1) = "," Then
'        getTagsProperty = Right$(getTagsProperty, Len(getTagsProperty) - 1)
'    End If
'
'End Function

Sub MNavigate(sUrl As String, IE As WebBrowser, Optional ByVal Frame As String)

Dim sProtocol As String
Dim sMain As String
Dim sSec As String
Dim sExt As String


sProtocol = linvblib.LeftLeft(sUrl, ":")
If Len(sProtocol) = 1 Then

    sMain = linvblib.LeftLeft(sUrl, "|/", vbBinaryCompare, ReturnOriginalStr)
    sSec = linvblib.LeftRight(sUrl, "|/", vbBinaryCompare, ReturnEmptyStr)
    sExt = LCase$(linvblib.RightRight(sMain, ".", vbBinaryCompare, ReturnEmptyStr))
    If sExt = "zip" Or sExt = "zhtm" Or sExt = "zjpg" Then
        MainFrm.loadzh sMain, sSec
    End If
Else
    IE.Navigate sUrl, , Frame
End If
End Sub
