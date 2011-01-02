VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPictureGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EErrorPictureGlass
    eeBasePictureGlass = 13140   ' CPictureGlass
    eePictureNotBitmap           ' Picture must contain bitmap
    eeInvalidCanvas              ' Drawing surface lacks required properties
End Enum

Private cvsDst As Object, hdcDst As Long, clrMask As Long
Private hdcImage As Long, hbmpImage As Long, hbmpImageOld As Long
Private hdcMask As Long, hbmpMask As Long, hbmpMaskOld As Long
Private hdcBack As Long, hbmpBack As Long, hbmpBackOld As Long
Private hdcCache As Long, hbmpCache As Long, hbmpCacheOld As Long
Private fExist As Boolean, fVisible As Boolean
Private xOld As Long, yOld As Long
Private dxSrc As Long, dySrc As Long
Private xLeft As Long, yTop As Long

Sub Create(cvsDstA As Object, picSrc As Picture, clrMaskA As Long, _
           Optional x As Variant, Optional y As Variant)
    
    ' Clean up any old instance before creating a new one
    If fExist Then Destroy
    ' Save at module level for use in properties and methods
    clrMask = clrMaskA
    Set cvsDst = cvsDstA
    If picSrc.Type <> vbPicTypeBitmap Then ErrRaise eePictureNotBitmap
    
    ' Catch any errors from canvas that doesn't have needed properties
    On Error GoTo CreateErrorCanvas
    With cvsDst
        hdcDst = .hDC
        ' Get size and position of image in pixels
        dxSrc = .ScaleX(picSrc.Width, vbHimetric, vbPixels)
        dySrc = .ScaleY(picSrc.Height, vbHimetric, vbPixels)
        ' Default is the center
        If IsMissing(x) Then x = .ScaleWidth / 2
        If IsMissing(y) Then y = .ScaleHeight / 2
        xLeft = .ScaleX(x, .ScaleMode, vbPixels)
        yTop = .ScaleY(y, .ScaleMode, vbPixels)
    End With
    Dim cPlanes As Long, cPixelBits As Long
    cPlanes = GetDeviceCaps(hdcDst, PLANES)
    cPixelBits = GetDeviceCaps(hdcDst, BITSPIXEL)
    
    ' Create memory DC compatible with screen for picture copy
    Dim hdcSrc As Long, hdcSrcOld As Long, hbmpSrcOld As Long
    hdcSrc = CreateCompatibleDC(0&)
    ' Select bitmap into DC
    hbmpSrcOld = SelectObject(hdcSrc, picSrc.Handle)
    
    ' Create memory DC for image with inverted background (AND mask)
    hdcImage = CreateCompatibleDC(0&)
    ' Create color bitmap same as screen
    hbmpImage = CreateBitmap(dxSrc, dySrc, cPlanes, cPixelBits, 0&)
    hbmpImageOld = SelectObject(hdcImage, hbmpImage)
    ' Make copy of picture because we don't want to modify original
    Call BitBlt(hdcImage, 0, 0, dxSrc, dySrc, hdcSrc, 0, 0, vbSrcCopy)
    
    ' Create DC for monochrome mask of image (XOR mask)
    hdcMask = CreateCompatibleDC(0&)
    ' Create bitmap (monochrome by default)
    hbmpMask = CreateCompatibleBitmap(hdcMask, dxSrc, dySrc)
    ' Select it into DC
    hbmpMaskOld = SelectObject(hdcMask, hbmpMask)
    ' Set background of source to the mask color
    Call SetBkColor(hdcSrc, clrMask)
    ' Copy color bitmap to monochrome DC to create mono mask
    Call BitBlt(hdcMask, 0, 0, dxSrc, dySrc, hdcSrc, 0, 0, vbSrcCopy)
    
    ' We've copied and used the source picture, so give it back
    Call SelectObject(hdcSrc, hbmpSrcOld)
    Call DeleteDC(hdcSrc)
    
    ' Invert background of image to create AND Mask
    Call SetBkColor(hdcImage, vbBlack)
    Call SetTextColor(hdcImage, vbWhite)
    Call BitBlt(hdcImage, 0, 0, dxSrc, dySrc, hdcMask, 0, 0, vbSrcAnd)
                
    ' Create memory DCs for old background and cache
    hdcBack = CreateCompatibleDC(0&)
    hbmpBack = CreateBitmap(dxSrc, dySrc, cPlanes, cPixelBits, 0&)
    hbmpBackOld = SelectObject(hdcBack, hbmpBack)
    hdcCache = CreateCompatibleDC(0&)
    hbmpCache = CreateBitmap(dxSrc, dySrc, cPlanes, cPixelBits, 0&)
    hbmpCacheOld = SelectObject(hdcCache, hbmpCache)
    
    ' Invalid x and y indicate first move hasn't occurred
    xOld = -1: yOld = -1
    fExist = True: fVisible = True
    Exit Sub
CreateErrorCanvas:
    ErrRaise eeInvalidCanvas
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Sub Destroy()
    BugAssert fExist
    ' Select old mask back to DC
    Call SelectObject(hdcMask, hbmpMaskOld)
    ' Now it's safe to delete DC and bitmask
    Call DeleteDC(hdcMask)
    Call DeleteObject(hbmpMask)
    ' Clean up inverted image DC
    Call SelectObject(hdcImage, hbmpImageOld)
    Call DeleteDC(hdcImage)
    Call DeleteObject(hbmpImage)
    ' Clean up cache DC
    Call SelectObject(hdcCache, hbmpCacheOld)
    Call DeleteDC(hdcCache)
    Call DeleteObject(hbmpCache)
    ' Clean up old background DC
    Call SelectObject(hdcBack, hbmpBackOld)
    Call DeleteDC(hdcBack)
    Call DeleteObject(hbmpBack)
    xOld = -1: yOld = -1
    fExist = False
End Sub

Public Sub Draw()
With cvsDst
    BugAssert fExist
    If fVisible = False Then Exit Sub
    
    ' Copy old background to its last location
    If xOld <> -1 Then
        Call BitBlt(hdcDst, xOld, yOld, dxSrc, dySrc, _
                    hdcBack, 0, 0, vbSrcCopy)
    End If
    ' Save current background and position for next time
    Call BitBlt(hdcBack, 0, 0, dxSrc, dySrc, _
                hdcDst, xLeft, yTop, vbSrcCopy)
    ' Create cache copy of background to work on
    Call BitBlt(hdcCache, 0, 0, dxSrc, dySrc, _
                hdcDst, xLeft, yTop, vbSrcCopy)
    xOld = xLeft: yOld = yTop
    ' Save color and set to white and black
    Dim clrBack As Long, clrFore As Long
    clrBack = GetBkColor(hdcCache)
    clrFore = GetTextColor(hdcCache)
    Call SetBkColor(hdcCache, vbWhite)
    Call SetTextColor(hdcCache, vbBlack)
    ' Mask the background
    Call BitBlt(hdcCache, 0, 0, dxSrc, dySrc, hdcMask, 0, 0, vbSrcAnd)
    ' Put image in hole created by mask
    Call BitBlt(hdcCache, 0, 0, dxSrc, dySrc, hdcImage, 0, 0, vbSrcPaint)
    ' Restore color
    Call SetBkColor(hdcCache, clrBack)
    Call SetTextColor(hdcCache, clrFore)
    ' Put finished cache on screen
    Call BitBlt(hdcDst, xLeft, yTop, dxSrc, dySrc, _
                hdcCache, 0, 0, vbSrcCopy)
End With
End Sub

Public Sub Remove()
    BugAssert fExist
    If fVisible = False Then Exit Sub
        
    ' Copy the old background to its last location
    If xOld <> -1 Then
        Call BitBlt(hdcDst, xOld, yOld, dxSrc, dySrc, _
                    hdcBack, 0, 0, vbSrcCopy)
    End If
End Sub

Public Sub Move(xLeftA As Long, Optional yTopA As Long = -1)
With cvsDst
    BugAssert fExist
    xLeft = .ScaleX(xLeftA, .ScaleMode, vbPixels)
    If yTopA <> -1 Then yTop = .ScaleY(yTopA, .ScaleMode, vbPixels)
    Draw
End With
End Sub

Property Get MaskColor() As Long
    BugAssert fExist
    MaskColor = clrMask
End Property

Property Get Left() As Single
    BugAssert fExist
    Left = cvsDst.ScaleX(xLeft, vbPixels, cvsDst.ScaleMode)
End Property

Property Let Left(rLeft As Single)
    BugAssert fExist
    xLeft = cvsDst.ScaleX(rLeft, cvsDst.ScaleMode, vbPixels)
    Draw
End Property

Property Get Top() As Single
    BugAssert fExist
    Top = cvsDst.ScaleY(yTop, vbPixels, cvsDst.ScaleMode)
End Property

Property Let Top(rTop As Single)
    BugAssert fExist
    yTop = cvsDst.ScaleY(rTop, cvsDst.ScaleMode, vbPixels)
    Draw
End Property

Property Get Width() As Single
    BugAssert fExist
    Width = cvsDst.ScaleX(dxSrc, vbPixels, cvsDst.ScaleMode)
End Property

Property Get Height() As Single
    BugAssert fExist
    Height = cvsDst.ScaleY(dySrc, vbPixels, cvsDst.ScaleMode)
End Property

Property Get Visible() As Boolean
    BugAssert fExist
    Visible = fVisible
End Property

Property Let Visible(fVisibleA As Boolean)
    BugAssert fExist
    fVisible = fVisibleA
End Property
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".PictureGlass"
        Select Case e
        Case eeBasePictureGlass
            BugAssert True
        Case eePictureNotBitmap
            sText = "Picture must contain bitmap"
        Case eeInvalidCanvas
            sText = "Drawing surface lacks required properties"
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If


