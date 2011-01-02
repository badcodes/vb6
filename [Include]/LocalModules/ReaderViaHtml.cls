VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReaderViaHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private zhReaderTemp As New clsTempFile
'Private bReallyComplete As Boolean
'Const NavigateTimeOut As Integer = 10 ' second
Implements IReader
Private Sub Class_Initialize()
Set IReader = Me
End Sub
Private Sub IReader_eBeforeNavigate(ByRef URL As Variant, ByRef Cancel As Boolean, ByRef IEView As WebBrowser, Optional targetFrameName As String)
    'zhrStatus.sCur_zhSubFile = ""

    Dim sBaseDir As String
    Dim sLocalUrl As String
    Dim thefile As String
    Dim sFilesToUnzip As String
    Dim stmp As String

    If linvblib.PathExists(zhrStatus.sCur_zhFile) = False Then Exit Sub

    sLocalUrl = toUnixPath(CStr(URL))
    sBaseDir = sTempZH

    If InStr(1, LCase$(sLocalUrl), LCase$(sBaseDir), vbTextCompare) <> 1 Then Exit Sub

    If Left$(sLocalUrl, 5) = "file:" And Len(sLocalUrl) > 7 Then sLocalUrl = Right$(sLocalUrl, Len(sLocalUrl) - 8)
    thefile = Right$(sLocalUrl, Len(sLocalUrl) - Len(sBaseDir) - 1)

    If linvblib.PathExists(sLocalUrl) = False Then
        MainFrm.myXUnzip zhrStatus.sCur_zhFile, thefile, sTempZH, zhrStatus.sPWD
    End If

    If linvblib.PathExists(sLocalUrl) = False Then Exit Sub

    If chkFileType(thefile) <> ftIE Then
        Cancel = True
        MGetView thefile, IEView, targetFrameName
        Exit Sub
    End If

    MainFrm.LeftFrame.Enabled = False

    zhrStatus.sCur_zhSubFile = thefile

    If Right$(thefile, Len(TempHtm)) = TempHtm Then
        zhrStatus.sCur_zhSubFile = Replace(thefile, TempHtm, "")
        Exit Sub
    End If
    

    Debug.Print "getFilenameInQuotes start :" & Timer
    stmp = getFilenameInQuotes(sLocalUrl, sBaseDir)
    Debug.Print "getFilenameInQuotes end :" & Timer
    
    If stmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & stmp
    
    Debug.Print "getTagsProperty start :" & Timer
    
    Dim htmlHandle As New CHtmLPropHelper
    If htmlHandle.InitFrom(sLocalUrl) Then
    
        stmp = getTagsProperty(htmlHandle, sLocalUrl, "", "src", sBaseDir)
        If stmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & stmp
        
        stmp = getTagsProperty(htmlHandle, sLocalUrl, "", "background", sBaseDir)
        If stmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & stmp
        
        stmp = getTagsProperty(htmlHandle, sLocalUrl, "link", "href", sBaseDir)
        If stmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & stmp

    End If
    
    Set htmlHandle = Nothing
    
    Debug.Print "getTagsProperty end." & Timer
'    sTmp = getTagsProperty(sLocalUrl, "embed", "src", sBaseDir)
'    If sTmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & sTmp
'
'    sTmp = getTagsProperty(sLocalUrl, "link", "href", sBaseDir)
'
'    If sTmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & sTmp
'
'    sTmp = getTagsProperty(sLocalUrl, "body", "background", sBaseDir)
'
'    If sTmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & sTmp
'
'    sTmp = getTagsProperty(sLocalUrl, "table", "background", sBaseDir)
'
'    If sTmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & sTmp
'
'    sTmp = getTagsProperty(sLocalUrl, "td", "background", sBaseDir)
'
'    If sTmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & sTmp
'
'    sTmp = getTagsProperty(sLocalUrl, "img", "src", sBaseDir)
'
'    If sTmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & sTmp
'
'    sTmp = getTagsProperty(sLocalUrl, "script", "src", sBaseDir)
'
'    If sTmp <> "" Then sFilesToUnzip = sFilesToUnzip & "|" & sTmp

    If Left$(sFilesToUnzip, 1) = "|" Then sFilesToUnzip = Right$(sFilesToUnzip, Len(sFilesToUnzip) - 1)

    If sFilesToUnzip = "" Then Exit Sub
    
    MainFrm.myXUnzip zhrStatus.sCur_zhFile, sFilesToUnzip, sTempZH, zhrStatus.sPWD
Errorexit:

    MainFrm.Enabled = True

End Sub

Private Sub IReader_eNavigateComplete(ByRef URL As Variant, ByRef IEView As WebBrowser)

Dim sUrl As String, sTemp As String

sUrl = linvblib.toUnixPath(CStr(URL))
sTemp = linvblib.toUnixPath(sTempZH)
If sTemp = "" Then Exit Sub
If Left$(sUrl, Len(sTemp)) = sTemp Then
    sUrl = Right$(sUrl, Len(sUrl) - Len(sTemp))
    sUrl = zhrStatus.sCur_zhFile & "|" & sUrl
    If Right$(sUrl, Len(TempHtm)) = TempHtm Then sUrl = Left$(sUrl, Len(sUrl) - Len(TempHtm))
    MainFrm.AddUniqueItem MainFrm.cmbAddress, zhrStatus.sCur_zhFile '.cmbAddress, sUrl
    MainFrm.cmbAddress.text = sUrl
Else
    'MainFrm.AddUniqueItem MainFrm.cmbAddress, CStr(URL)
    MainFrm.cmbAddress.text = CStr(URL)
End If


'MainFrm.AddUniqueItem cmbAddress, cmbAddress.text

   ' MainFrm.LeftFrame.Enabled = True
    'MainFrm.StsBar.Panels("reading").text = zhrStatus.sCur_zhSubFile

End Sub

Private Sub IReader_endUP()

End Sub

Private Sub IReader_eStatusTextChange(ByVal text As String, ByRef IEView As WebBrowser)

    Dim sSText As String
    Dim sVText As String

    If zhrStatus.sCur_zhFile = "" Then
        Exit Sub
    Else
        sSText = Replace(text, "/", "\")
        sSText = Replace(sSText, "%20", " ")
        sVText = "file:\\\" & Replace(sTempZH, "/", "\") & "\"
        sSText = UCase$(sSText)
        sVText = UCase$(sVText)
        'MDebug.DPrint sSText
        'MDebug.DPrint sVText
        If InStr(sSText, sVText) Then
            MainFrm.StsBar.Panels("ie").text = Replace(sSText, sVText, zhrStatus.sCur_zhFile & "|\")
        Else
            MainFrm.StsBar.Panels("ie").text = text
        End If

        'MainFrm.StsBar.SimpleText = zhrstatus.sCur_zhFile  Replace(text, sHttpServer, "")
    End If
    



End Sub

Private Sub IReader_MGetView(shortfile As String, ByRef IEView As WebBrowser, Optional ByRef targetFrameName As String)

    If shortfile = "" Then MainFrm.appHtmlAbout: Exit Sub
    Dim fso As New gCFileSystem
    Dim tempfile As String
    Dim tempFile2 As String
    Dim bUseTemplate As Boolean
    Dim sTemplateFile As String
    Dim mapPoint As String
    tempfile = linvblib.RightLeft(shortfile, "#", vbTextCompare, ReturnOriginalStr)
    mapPoint = linvblib.RightRight(shortfile, "#", vbTextCompare, ReturnEmptyStr)
    tempfile = fso.BuildPath(sTempZH, tempfile)
    If fso.PathExists(tempfile) = False Then
        MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, sTempZH, zhrStatus.sPWD
    End If
    If fso.PathExists(tempfile) = False Then Exit Sub
    tempFile2 = tempfile & TempHtm
    
    sTemplateFile = MainFrm.IEView.Tag 'iniGetSetting(zhtmini, "Viewstyle", "TemplateFile")
    bUseTemplate = (Val(MainFrm.Tag) <> 0) 'CBoolStr(iniGetSetting(zhtmini, "ViewStyle", "UseTemplate"))

    MainFrm.NotPreOperate = True
    
    zhrStatus.sCur_zhSubFile = shortfile
    Select Case chkFileType(tempfile) 'file.bas

    Case ftIE
        MainFrm.NotPreOperate = False
        If mapPoint = "" Then
            IEView.Navigate2 tempfile, , targetFrameName
        Else
            IEView.Navigate2 tempfile & "#" & mapPoint, , targetFrameName
        End If
        Exit Sub
    Case ftZIP
        MainFrm.loadzh tempfile
    Case ftZhtm
        MainFrm.loadzh tempfile
    Case ftIMG
        Dim imgToLoad As Picture
        Dim imgHeight As Long
        Dim imgWidth As Long
        Dim screenHeight As Long
        Dim screenWidth As Long
        Dim resizeRateY As Double
        Dim resizeRateX As Double
        Dim resizeRate As Double
        Set imgToLoad = LoadPicture(tempfile)
        imgHeight = MainFrm.ScaleX(imgToLoad.Height, 8, 3)
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
            If createHtmlFromTemplate(tempfile, sTemplateFile, tempFile2, imgHeight, imgWidth) Then
                IEView.Navigate2 tempFile2, , targetFrameName
            ElseIf createDefaultHtml(tempfile, tempFile2, imgHeight, imgWidth) Then
                IEView.Navigate2 tempFile2, , targetFrameName
            Else
                IEView.Navigate2 tempfile, , targetFrameName
            End If
        ElseIf createDefaultHtml(tempfile, tempFile2, imgHeight, imgWidth) Then
            IEView.Navigate2 tempFile2, , targetFrameName
        Else
            IEView.Navigate2 tempfile, , targetFrameName
        End If
    Case Else

        If bUseTemplate Then

            If createHtmlFromTemplate(tempfile, sTemplateFile, tempFile2) Then
                IEView.Navigate2 tempFile2, , targetFrameName
            ElseIf createDefaultHtml(tempfile, tempFile2) Then
                IEView.Navigate2 tempFile2, , targetFrameName
            Else
                IEView.Navigate2 tempfile, , targetFrameName
            End If

        ElseIf createDefaultHtml(tempfile, tempFile2) Then
            IEView.Navigate2 tempFile2, , targetFrameName
        Else
            IEView.Navigate2 tempfile, , targetFrameName
        End If

    End Select

End Sub


Public Function getFilenameInQuotes(ByRef htmlfile As String, ByVal sBaseDir As String) As String
Dim lCount As Long, lLoop As Long
Dim stmp As String
Dim pValue() As String
Dim fNum As Integer


Dim fso As New FileSystemObject
If fso.FileExists(htmlfile) = False Then Exit Function

fNum = FreeFile()
Open htmlfile For Input As #fNum
lCount = MClassicIO.StrBetween(fNum, Chr$(34), Chr$(34), pValue, , False)
Close #fNum

If lCount < 1 Then Exit Function
For lLoop = 1 To lCount
    stmp = pValue(lLoop)
    If linvblib.gCString.isTypicalFileName(stmp) = False Then GoTo forContinue
    If InStr(stmp, ":") > 0 Then GoTo forContinue
    stmp = fso.BuildPath(fso.GetParentFolderName(htmlfile), stmp)
    stmp = fso.GetAbsolutePathName(stmp)
    stmp = toUnixPath(stmp)
    If InStr(1, stmp, sBaseDir, vbTextCompare) = 1 Then
        stmp = Right$(stmp, Len(stmp) - Len(sBaseDir) - 1)
        getFilenameInQuotes = getFilenameInQuotes & "|" & stmp
    End If
forContinue:
Next
If Left$(getFilenameInQuotes, 1) = "|" Then
    getFilenameInQuotes = Right$(getFilenameInQuotes, Len(getFilenameInQuotes) - 1)
End If
End Function
Public Function getTagsProperty(ByRef htmlHandle As CHtmLPropHelper, ByRef htmlfile As String, ByVal tagName As String, ByVal propertyName As String, ByVal sBaseDir As String) As String


    Dim fso As New FileSystemObject
    Dim lCount As Long
    Dim lLoop As Long
    Dim sPropertyValue As String
    Dim stmp As String
    Dim pValue() As String
    
    If fso.FileExists(htmlfile) = False Then Exit Function
        
    lCount = htmlHandle.GetPropertys(pValue(), propertyName, tagName)

    If lCount < 1 Then Exit Function

    For lLoop = 1 To lCount
        sPropertyValue = pValue(lLoop)
        If sPropertyValue = "" Then GoTo forContinue
        If InStr(sPropertyValue, "%") > 0 Then
        sPropertyValue = linvblib.UnescapeUrl(sPropertyValue)
        End If
        stmp = sPropertyValue

        If InStr(stmp, ":") > 0 Then GoTo forContinue

        stmp = fso.BuildPath(fso.GetParentFolderName(htmlfile), stmp)
        stmp = fso.GetAbsolutePathName(stmp)
        stmp = toUnixPath(stmp)
        'sBaseDir = replaceSlash(sBaseDir)

        If InStr(1, stmp, sBaseDir, vbTextCompare) = 1 Then
            stmp = Right$(stmp, Len(stmp) - Len(sBaseDir) - 1)
            getTagsProperty = getTagsProperty & "|" & stmp
        End If

forContinue:
    Next

    If Left$(getTagsProperty, 1) = "|" Then
        getTagsProperty = Right$(getTagsProperty, Len(getTagsProperty) - 1)
    End If

End Function


Private Sub IReader_MNavigate(sUrl As String, IE As WebBrowser, Optional targetFrameName As String)

Dim sProtocol As String
Dim sMain As String
Dim sSec As String
Dim sExt As String
'Dim iAsc As Integer

sProtocol = UCase$(linvblib.LeftLeft(sUrl, ":"))

If Len(sProtocol) = 1 And Asc(sProtocol) > 64 And Asc(sProtocol) < 91 Then
    sMain = linvblib.LeftLeft(sUrl, "|/", vbBinaryCompare, ReturnOriginalStr)
    sSec = linvblib.LeftRight(sUrl, "|/", vbBinaryCompare, ReturnEmptyStr)
    sExt = LCase$(linvblib.RightRight(sMain, ".", vbBinaryCompare, ReturnEmptyStr))
    If sExt = "zip" Or sExt = "zhtm" Or sExt = "zjpg" Then
        MainFrm.loadzh sMain, sSec
    End If
Else
    IE.Navigate sUrl, , targetFrameName
End If
End Sub

Private Sub IReader_startUP()
End Sub
