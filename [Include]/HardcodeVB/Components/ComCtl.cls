VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCommonControl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorCommonControl
    eeBaseCommonControl = 13440     ' CommonControl
End Enum

' Helpers for common control functions and image lists

Function INDEXTOOVERLAYMASK(i As Long) As Long
    INDEXTOOVERLAYMASK = i * 256
End Function

' ImageList_ReplaceIcon(himl, -1, hicon)
Function ImageList_AddIcon(ByVal himl As Long, ByVal hIcon As Long) As Long
    ImageList_AddIcon = ImageList_ReplaceIcon(himl, -1, hIcon)
End Function

' ImageList_Remove(himl, -1)
Function ImageList_RemoveAll(ByVal himl As Long) As Long
    ImageList_RemoveAll = ImageList_Remove(himl, -1)
End Function

' ImageList_GetIcon(himl, i, 0)
Function ImageList_ExtractIcon(ByVal himl As Long, ByVal i As Long) As Long
    ImageList_ExtractIcon = ImageList_GetIcon(himl, i, 0)
End Function

' ImageList_LoadImage(hi, lpbmp, cx, cGrow, crMask, IMAGE_BITMAP, 0)
Function ImageList_LoadBitmap(ByVal hi As Long, ByVal lpbmp As String, _
    ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, _
    ByVal uType As Long, ByVal uFlags As Long) As Long
    ImageList_LoadBitmap = ImageList_LoadImage(hi, lpbmp, cx, cGrow, _
                                               crMask, IMAGE_BITMAP, 0)
End Function

#If fComponent Then
Sub DrawImage(imlst As Object, vIndex As Variant, ByVal hDC As Long, _
              ByVal x As Long, ByVal y As Long, _
              Optional ByVal afDraw As EILD = ILD_TRANSPARENT)
#Else
Sub DrawImage(imlst As Control, vIndex As Variant, ByVal hDC As Long, _
              ByVal x As Long, ByVal y As Long, _
              Optional ByVal afDraw As EILD = ILD_TRANSPARENT)
#End If
    ImageList_Draw imlst.hImageList, _
                   imlst.ListImages(vIndex).Index - 1, hDC, _
                   x / Screen.TwipsPerPixelX, _
                   y / Screen.TwipsPerPixelY, afDraw
End Sub


' System image lists

Function GetSysImageList(cCount As Long, _
                         Optional ByVal fLargeIcon As Boolean = True) As Long
    Dim shfi As SHFILEINFO
    Dim hSysIm As Long, hIcon As Long, af As Long
    af = SHGFI_SYSICONINDEX Or _
         IIf(fLargeIcon, SHGFI_LARGEICON, SHGFI_SMALLICON)
    hSysIm = SHGetFileInfo(Left$(CurDir$, 3), 0, shfi, Len(shfi), af)
    cCount = ImageList_GetImageCount(hSysIm)
    GetSysImageList = hSysIm
End Function

Function GetSysIcon(ByVal hSysIm As Long, ByVal i As Integer, _
                    Optional ByVal xWidth As Long, _
                    Optional ByVal yHeight As Long) As Picture
    Set GetSysIcon = Nothing
    Dim f As Boolean, cx As Long, cy As Long
    Dim iminf As IMAGEINFO
    f = ImageList_GetImageInfo(hSysIm, i, iminf)
    If Not f Then Exit Function
    f = ImageList_GetIconSize(hSysIm, cx, cy)
    If Not f Then Exit Function
    ' These just go to temporary variables if missing
    xWidth = cx
    yHeight = cx
    ' Check for bitmap
    If iminf.hbmMask = hNull Then Exit Function
    ' Get icon handle and convert to picture
    Set GetSysIcon = MPicTool.IconToPicture(ImageList_GetIcon(hSysIm, i, ILD_NORMAL))
    
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".CommonControl"
        Select Case e
        Case eeBaseCommonControl
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

