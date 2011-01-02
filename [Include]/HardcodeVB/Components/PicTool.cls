VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GPicTool"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub OleCreatePictureIndirect Lib "olepro32.dll" ( _
    lpPictDesc As PICTDESC, riid As UUID, _
    ByVal fPictureOwnsHandle As Long, ipic As IPicture)

Private dxyShell As Long

Private Declare Function GetObjectBitmap Lib "GDI32" Alias "GetObjectA" ( _
    ByVal hBitmap As Long, ByVal cbBuffer As Long, _
    lpBitmap As BITMAP) As Long
    
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long 'LPVOID
End Type

Public Enum EErrorPicTool
    eeBasePicTool = 13560   ' PicTool
End Enum

Public Enum EIconSize
    eisDefault = -1
    eisImage = -2
    eisSmall = -3
    eisHuge = -4
    eisShell = -5
End Enum

Public Enum EConversions
    TwipsPerPoint = 20
    TwipsPerCharX = 120
    TwipsPerCharY = 240
    TwipsPerInch = 1440
    TwipsPerDecimeter = 5669
End Enum

Private iidIPicture As UUID
#If fComponent = 0 Then
Private fInitialized As Boolean
#End If

Sub Class_Initialize()
    ' Initialize iidIPicture GUID constant from string
    IIDFromString "{7BF80980-BF32-101A-8BBB-00AA00300CAB}", iidIPicture
#If fComponent = 0 Then
    ' Set initialized flag for standard module
    fInitialized = True
#End If
End Sub

Private Sub InitIf()
#If fComponent = 0 Then
    If Not fInitialized Then Class_Initialize
#End If
End Sub

' Scale conversion procedures

Function TwipsPerCentimeter() As Single
    TwipsPerCentimeter = 566.9
End Function

Function TwipsPerMillimeter() As Single
    TwipsPerMillimeter = 56.69
End Function

Function TwipsPerHiMetricUnit() As Single
    TwipsPerHiMetricUnit = 0.5669
End Function

Function PicXToPixel(ByVal xHiMetric As Long) As Long
    PicXToPixel = xHiMetric * TwipsPerDecimeter / Screen.TwipsPerPixelX / 10000
End Function

Function PicYToPixel(ByVal yHiMetric As Long) As Long
    PicYToPixel = yHiMetric * TwipsPerDecimeter / Screen.TwipsPerPixelY / 10000
End Function

'' Picture conversion procedures

Function IconToPicture(ByVal hIcon As Long) As IPicture
    If hIcon = hNull Then Exit Function
    Dim ipic As IPicture, picdes As PICTDESC
    ' Fill picture description
    picdes.cbSizeofstruct = Len(picdes)
    picdes.picType = vbPicTypeIcon
    picdes.hgdiobj = hIcon
    
    InitIf  ' Initialize picture GUID if necessary
    
    ' Create picture from icon handle
    OleCreatePictureIndirect picdes, iidIPicture, True, ipic
    ' Result will be valid Picture or Nothing--either way set it
    Set IconToPicture = ipic
End Function

Function CursorToPicture(ByVal hIcon As Long) As IPicture
    ' It's just an alias
    Set CursorToPicture = IconToPicture(hIcon)
End Function

Function BitmapToPicture(ByVal hBmp As Long, _
                         Optional ByVal hPal As Long = hNull) _
                         As IPicture
    ' Fill picture description
    Dim ipic As IPicture, picdes As PICTDESC
    picdes.cbSizeofstruct = Len(picdes)
    picdes.picType = vbPicTypeBitmap
    picdes.hgdiobj = hBmp
    picdes.hPalOrXYExt = hPal
    
    InitIf  ' Initialize picture GUID if necessary
    
    ' Create picture from bitmap handle
    OleCreatePictureIndirect picdes, iidIPicture, True, ipic
    ' Result will be valid Picture or Nothing--either way set it
    Set BitmapToPicture = ipic
End Function

Function MetafileToPicture(ByVal hMeta As Long, _
                           ByVal xExt As Integer, _
                           ByVal yExt As Integer, _
                           Optional fOld As Boolean) As IPicture
    If hMeta = hNull Then Exit Function
    Dim ipic As IPicture, picdes As PICTDESC
    ' Fill picture description (assume enhanced)
    picdes.cbSizeofstruct = Len(picdes)
    If fOld Then
        picdes.picType = vbPicTypeMetafile
    Else
        picdes.picType = vbPicTypeEMetafile
    End If
    picdes.hgdiobj = hMeta
    picdes.hPalOrXYExt = MBytes.MakeDWord(xExt, yExt) ' Fake union
    
    InitIf  ' Initialize picture GUID if necessary
    ' Create picture from icon handle
    OleCreatePictureIndirect picdes, iidIPicture, True, ipic
    ' Result will be valid Picture or Nothing--either way set it
    Set MetafileToPicture = ipic
End Function

' Create a mask on destination DC from source DC of specified size
Function MakeMask(picSrc As StdPicture) As StdPicture
    Dim hdcSrc As Long, hbmpSrc As Long
    Dim hdcDst As Long, hbmpDst As Long
    Dim dxSrc As Long, dySrc As Long
    
    ' Get picture size
    dxSrc = PicXToPixel(picSrc.Width)
    dySrc = PicYToPixel(picSrc.Height)
    
    ' Select source into memory DC
    
    
    ' Create memory device context for destination
    hdcDst = CreateCompatibleDC(0)
    ' Create monochrome bitmap and select it into DC
    hbmpDst = CreateCompatibleBitmap(hdcDst, dxSrc, dySrc)
    hbmpDst = SelectObject(hdcDst, hbmpDst)
    ' Copy color bitmap to DC to create mono mask
    BitBlt hdcDst, 0, 0, dxSrc, dySrc, hdcSrc, 0, 0, SRCCOPY
    ' Clean up
    Call SelectObject(hdcDst, hbmpDst)
    Call DeleteObject(hbmpDst)
    Call DeleteDC(hdcDst)
    
    'Set MakeMask = BitmapToPicture(hbmpDst)
End Function

'' Handle information procedures

Sub GetIconSize(ByVal hIcon As Long, dx As Long, dy As Long, _
                Optional xHot As Long, Optional yHot As Long)
    Dim ico As ICONINFO, bmp As BITMAP, dc As Long, f As Boolean
    f = GetIconInfo(hIcon, ico)
    f = GetObjectBitmap(ico.hbmColor, LenB(bmp), bmp)
    dx = bmp.bmWidth
    dy = bmp.bmHeight
    xHot = ico.xHotspot
    yHot = ico.yHotspot
End Sub

Sub GetBitmapSize(ByVal hBitmap As Long, dx As Long, dy As Long)
    Dim bmp As BITMAP, f As Boolean
    f = GetObjectBitmap(hBitmap, LenB(bmp), bmp)
    dx = bmp.bmWidth
    dy = bmp.bmHeight
End Sub

Function GetShellIconSize() As Long
#If 1 Then
    ' Grabbing size out of registry works, but might change
    Const sMetrics = "Control Panel\Desktop\WindowMetrics"
    On Error Resume Next
    GetShellIconSize = MRegTool.GetRegStr(sMetrics, "Shell Icon Size")
    ' If size isn't in registry, assume the default size
    If Err Then GetShellIconSize = 32
#Else
    ' Recommended way of getting size doesn't work until after login
    Dim hImlst As Long, fi As SHFILEINFO, cx As Long, cy As Long
    hImlst = SHGetFileInfo(".", 0, fi, Len(fi), _
                           SHGFI_SYSICONINDEX Or SHGFI_SHELLICONSIZE)
    If ImageList_GetIconSize(hImlst, cx, cy) Then
        GetShellIconSize = cx
    Else
        GetShellIconSize = -1
    End If
#End If
End Function

'' Resource helpers

Function ResourceIdToStr(ByVal ID As Long) As String
    Select Case ID
    Case RT_CURSOR
        ResourceIdToStr = "CURSOR"
    Case RT_BITMAP
        ResourceIdToStr = "BITMAP"
    Case RT_ICON
        ResourceIdToStr = "ICON"
    Case RT_MENU
        ResourceIdToStr = "MENU"
    Case RT_DIALOG
        ResourceIdToStr = "DIALOG"
    Case RT_STRING
        ResourceIdToStr = "STRING"
    Case RT_FONTDIR
        ResourceIdToStr = "FONTDIR"
    Case RT_FONT
        ResourceIdToStr = "FONT"
    Case RT_ACCELERATOR
        ResourceIdToStr = "ACCELERATOR"
    Case RT_RCDATA
        ResourceIdToStr = "RCDATA"
    Case RT_MESSAGETABLE
        ResourceIdToStr = "MESSAGETABLE"
    Case RT_GROUP_CURSOR
        ResourceIdToStr = "GROUP_CURSOR"
    Case RT_GROUP_ICON
        ResourceIdToStr = "GROUP_ICON"
    Case RT_VERSION
        ResourceIdToStr = "VERSION"
    Case RT_DLGINCLUDE
        ResourceIdToStr = "DLGINCLUDE"
    Case RT_PLUGPLAY
        ResourceIdToStr = "PLUGPLAY"
    Case RT_VXD
        ResourceIdToStr = "VXD"
    Case Else
        ResourceIdToStr = "Unknown"
    End Select
End Function

' The Win32 UnlockResource function is a macro returning zero. Since you
' can't emulate this in a type library, this do-nothing function is
' provided here. Better yet, don't try to unlock resources.
Function UnlockResource(ByVal hResData As Long) As Long
    UnlockResource = 0
End Function

Function LoadAnyPicture(Optional sPicture As String, _
                        Optional eis As EIconSize = eisDefault _
                        ) As Picture
    Dim hIcon As Long, sExt As String, xy As Long, af As Long
    ' If no picture, return Nothing (clears picture)
    If sPicture = sEmpty Then Exit Function
    ' Use default LoadPicture for all except icons with argument
    sExt = MUtility.GetFileExt(sPicture)
    If UCase$(sExt) <> ".ICO" Or eis = -1 Then
        Set LoadAnyPicture = VB.LoadPicture(sPicture)
        Exit Function
    End If
    
    Select Case eis
    Case eisSmall
        xy = 16: af = LR_LOADFROMFILE
    Case eisHuge
        xy = 48: af = LR_LOADFROMFILE
    Case eisImage
        xy = 0: af = LR_LOADFROMFILE
    Case eisShell ' Get icon size from system
        xy = GetShellIconSize(): af = LR_LOADFROMFILE
    Case Is > 0   ' Use arbitrary specified size--72 by 72 or whatever
        xy = eis: af = LR_LOADFROMFILE
    Case Else     ' Includes eisDefault
        xy = 0: af = LR_LOADFROMFILE Or LR_DEFAULTSIZE
    End Select
    hIcon = LoadImage(0&, sPicture, IMAGE_ICON, xy, xy, af)
    ' If this fails, use original load
    If hIcon <> hNull Then
        Set LoadAnyPicture = IconToPicture(hIcon)
    Else
        Set LoadAnyPicture = VB.LoadPicture(sPicture)
    End If
End Function

Function LoadAnyResPicture(vRes As Variant, iResType As Integer, _
                           Optional eis As EIconSize = eisDefault _
                           ) As Picture
#If fComponent Then
    Dim hIcon As Long, sExt As String, xy As Long, af As Long
    ' Can't use LoadImage in environment--have to make do with default
    If Not MUtility.IsExe() Then
        If (eis = -1) Or (iResType <> vbResIcon) Then
            Set LoadAnyResPicture = VB.LoadResPicture(vRes, iResType)
            Exit Function
        End If
    End If
    
    Select Case eis
    Case eisSmall
        xy = 16: af = LR_LOADFROMFILE
    Case eisHuge
        xy = 48: af = LR_LOADFROMFILE
    Case eisImage
        xy = 0: af = LR_LOADFROMFILE
    Case eisShell   ' Get icon size from system
        xy = GetShellIconSize(): af = LR_LOADFROMFILE
    Case Is > 0     ' Use arbitrary specified size--72 by 72 or whatever
        xy = eis: af = LR_LOADFROMFILE
    Case Else       ' Includes eisDefault
        xy = 0: af = LR_LOADFROMFILE Or LR_DEFAULTSIZE
    End Select
    If TypeName(vRes) = "String" Then
        hIcon = LoadImage(App.hInstance, CStr(vRes), IMAGE_ICON, xy, xy, af)
    Else
        hIcon = LoadImage(App.hInstance, CLng(vRes), IMAGE_ICON, xy, xy, af)
    End If
    If hIcon <> hNull Then
        Set LoadAnyResPicture = IconToPicture(hIcon)
    Else
        Set LoadAnyResPicture = VB.LoadResPicture(vRes, iResType)
    End If
#Else
    BugAssert True, "Can't use from DLL--function looks for resource in DLL rather than app"
#End If
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".PicTool"
        Select Case e
        Case eeBasePicTool
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.EXEName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

