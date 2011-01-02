VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GPalTool"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ECycleDirection
    ecdCycleLeft
    ecdCycleRight
    ecdCycleIn
    ecdCycleOut
End Enum

Public Enum EErrorPalTool
    eeBasePalTool = 13540   ' PalTool
End Enum

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

Sub DrawPalette(cvsDst As Object, hPal As Long, _
                Optional ByVal dx As Long, Optional ByVal dy As Long, _
                Optional ByVal x As Long, Optional ByVal y As Long)
    If dx = 0 Then dx = cvsDst.ScaleWidth
    If dy = 0 Then dy = cvsDst.ScaleHeight
    If hPal = 0 Then Exit Sub
    
    Dim c As Long, cPal As Integer, ape() As PALETTEENTRY
    ' Get the size of the palette and dimension the array to it
    cPal = PalSize(hPal)
    ReDim ape(0 To cPal - 1) As PALETTEENTRY
    ' Fill the array with palette colors
    c = GetPaletteEntries(hPal, 0&, cPal, ape(0))
    BugAssert c = cPal
    Dim i As Long, dxBar As Long, clr As Long
    ' Draw the palette colors as bars proportional to size of canvas
    dxBar = dx / cPal
    For i = 0 To cPal - 2
        x = i * dxBar
        cvsDst.Line (x, y)-(x + dxBar, y + dy), _
                    PaletteColorFromEntry(ape(i)), BF
    Next
End Sub

' Attempt to get the palette colors of a bitmap from its handle
Function GetBitmapPalette(ByVal hBmp As Long, _
                          Optional ByVal hDC As Long = hInvalid) As Long
    Dim bmp As BITMAP, bmpi As BITMAPINFO256
    Dim cClrBits As Integer, f As Long
 
    ' Retrieve bitmap color format, width, and height
    f = GetObjectBitmap(hBmp, LenB(bmp), bmp)
 
    ' Convert the color format to a count of bits
    cClrBits = bmp.bmPlanes * bmp.bmBitsPixel
 
    ' Like VB, we only return 256-color palettes
    If cClrBits <> 8 Then Exit Function
 
    ' Initialize the fields in the BITMAPINFOHEADER structure.
    bmpi.bmiHeader.biSize = LenB(bmpi.bmiHeader)
    bmpi.bmiHeader.biWidth = bmp.bmWidth
    bmpi.bmiHeader.biHeight = bmp.bmHeight
    bmpi.bmiHeader.biPlanes = bmp.bmPlanes
    bmpi.bmiHeader.biBitCount = bmp.bmBitsPixel
    bmpi.bmiHeader.biCompression = BI_RGB
    bmpi.bmiHeader.biClrUsed = 0
    bmpi.bmiHeader.biClrImportant = 0
    ' Compute the number of bytes in the array of color
    ' indices and store the result in SizeImage field
    Dim cSizeImage As Long
    cSizeImage = (bmp.bmWidth + 7) / 8 * bmp.bmHeight * cClrBits
    bmpi.bmiHeader.biSizeImage = cSizeImage
 
    ' Retrieve the color table (RGBQUAD array) and the bits
    ' (array of palette indices) from the DIB
    Dim hDCTmp As Long, cColors As Integer
    If hDC = hInvalid Then
        hDCTmp = GetDC(0)
    Else
        hDCTmp = hDC
    End If
    f = GetDIBits(hDCTmp, hBmp, 0, bmp.bmHeight, ByVal 0, _
                  bmpi, DIB_RGB_COLORS)
    ' Free the DC
    If hDC = hInvalid Then f = ReleaseDC(0, hDCTmp)
                             
    Dim lpal As LOGPALETTE256, i As Long, c As Long
    cColors = bmpi.bmiHeader.biClrUsed
    ' Like VB, we only return 256-color palettes
    If cColors = 0 Then cColors = 256
      
    ' RGBQUAD used by GetDIBColorTable has different format from PALETTEENTRY
    ' used by LOGPALETTE, so can't use CopyMemory
    For i = 0 To cColors - 1
        ' Skip black colors
        If i >= 1 And i < cColors - 1 Then
            If bmpi.bmiColors(i).rgbRed = 0 And _
               bmpi.bmiColors(i).rgbGreen = 0 And _
               bmpi.bmiColors(i).rgbBlue = 0 Then
                GoTo ContinueFor
            End If
        End If
        ' Copy and translate colors
        lpal.palPalEntry(i).peRed = bmpi.bmiColors(i).rgbRed
        lpal.palPalEntry(i).peGreen = bmpi.bmiColors(i).rgbGreen
        lpal.palPalEntry(i).peBlue = bmpi.bmiColors(i).rgbBlue
        lpal.palPalEntry(i).peFlags = 0
        c = c + 1
ContinueFor:
    Next
    lpal.palNumEntries = IIf(c Mod 2, c - 1, c)
    Debug.Print "Colors: " & lpal.palNumEntries
    lpal.palVersion = &H300
 
    ' Create and return the palette
    GetBitmapPalette = CreatePalette(lpal)

End Function
 
' Load a bitmap and its palette from a resource
Function LoadBitmapPalette(ByVal hMod As Long, vResource As Variant, _
                           hPal As Long) As Long
                           
    ' Make null in case of failure
    Dim hBmp As Long
    hPal = hNull
    
    Dim hRes As Long, hmemRes As Long, cRes As Long
    Dim pRes As Long, abRes() As Byte
    If VarType(vResource) = vbString Then
        hBmp = LoadImage(hMod, CStr(vResource), IMAGE_BITMAP, _
                         0, 0, LR_CREATEDIBSECTION)
        hRes = FindResourceStrId(hMod, CStr(vResource), RT_BITMAP)
        
    Else
        hBmp = LoadImageID(hMod, CLng(vResource), IMAGE_BITMAP, _
                           0, 0, LR_CREATEDIBSECTION)
        hRes = FindResourceIdId(hMod, CLng(vResource), RT_BITMAP)
    End If
    ' If bitmap found, return it
    If hBmp = hNull Then Exit Function
    LoadBitmapPalette = hBmp
    BugAssert hRes <> hNull     ' Shouldn't fail here
    ' Allocate memory block, and get its size
    hmemRes = LoadResource(hMod, hRes)
    cRes = SizeofResource(hMod, hRes)
    ' Lock it to get pointer
    pRes = LockResource(hmemRes)
    
    Dim bmpi As BITMAPINFO256
    If cRes > LenB(bmpi) Then cRes = LenB(bmpi)
    ' Copy memory block to array
    CopyMemory bmpi, ByVal pRes, cRes
    ' Free resource (no need to unlock)
    Call FreeResource(hmemRes)
    
    Dim lpal As LOGPALETTE256, cColors As Long, cBits As Long, i As Long
    cColors = bmpi.bmiHeader.biClrUsed
    cBits = bmpi.bmiHeader.biBitCount
    ' Like VB, we only return 256-color palettes
    If cBits <> 8 Then Exit Function
    If cColors = 0 Then cColors = 256

    ' RGBQUAD in BITMAPINFO has different format from PALETTEENTRY
    ' in LOGPALETTE, so can't use CopyMemory
    For i = 0 To cColors - 1
        ' Copy and translate colors
        lpal.palPalEntry(i).peRed = bmpi.bmiColors(i).rgbRed
        lpal.palPalEntry(i).peGreen = bmpi.bmiColors(i).rgbGreen
        lpal.palPalEntry(i).peBlue = bmpi.bmiColors(i).rgbBlue
        lpal.palPalEntry(i).peFlags = 0
    Next
    lpal.palNumEntries = cColors
    lpal.palVersion = &H300
 
    ' Create and return the palette through a reference
    hPal = CreatePalette(lpal)

End Function
 
Function PalSize(ByVal hPal As Long) As Integer
    Dim c As Integer, res As Long
    res = GetObjectPaletteEntries(hPal, 2, c)
    PalSize = c
End Function

Private Function PaletteColorFromEntry(pe As PALETTEENTRY) As OLE_COLOR
    ' Copy color bytes, ignore flag byte
    CopyMemory PaletteColorFromEntry, pe, 3
End Function

Private Sub PaletteColorToEntry(pe As PALETTEENTRY, ByVal clr As OLE_COLOR)
    ' Copy color bytes, ignore flag byte
    CopyMemory pe, clr, 3
End Sub

' Potentially the most efficient palette is an identity palette--one that is
' the same size as the system palette and that has the same system colors in its
' first and last entries. The user-defined colors go in the middle. This function
' takes a normal palette and returns the handle of an equivalent identity palette.
' It also returns the position of the first non-system color.
Function MakeIdentityPalette(ByVal hPal As Long, iFirst As Long) As Long
Dim logpal As LOGPALETTE256
With logpal
    
    Dim iLast As Long, hDC As Long, cStatic As Long, cPal As Long
    hDC = GetDC(hNull)
    ' Get the size of the palette and dimension the array to it
    .palVersion = &H300
    .palNumEntries = 256
    cPal = PalSize(hPal)
    
    ' We only deal with SYSPAL_STATIC mode (the most common)
    If GetSystemPaletteUse(hDC) <> SYSPAL_STATIC Then Exit Function
    
    ' Get the twenty static colors into array and then
    ' fill the empty spaces with the color table
    cStatic = GetDeviceCaps(hDC, NUMCOLORS)
    ' Too many colors for identity palette
    If cPal > 256 - cStatic Then Exit Function
    
    ' Get the system palette into the palette
    Dim i As Long, f As Long
    f = GetSystemPaletteEntries(hDC, 0, 256, .palPalEntry(0))
    iFirst = (cStatic \ 2)
    iLast = iFirst + cPal
    
    ' Fill the middle of the array with palette colors
    f = GetPaletteEntries(hPal, iFirst, cPal, .palPalEntry(iFirst))
    
    ' Set the peFlags of the lower static colors to zero
    For i = 0 To iFirst - 1
        .palPalEntry(i).peFlags = 0
    Next
    
    ' Mark our entries as PC_RESERVED
    For i = i To iFirst + cPal
        .palPalEntry(i).peFlags = PC_RESERVED
    Next
    
    ' Mark any other as black PC_RESERVED
    For i = i To 255 - (cStatic \ 2)
        .palPalEntry(i).peFlags = PC_RESERVED
        .palPalEntry(i).peBlue = 0
        .palPalEntry(i).peGreen = 0
        .palPalEntry(i).peRed = 0
    Next
    
    ' Set the peFlags of the upper static colors to zero
    For i = i To 255
        .palPalEntry(i).peFlags = 0
    Next
      
    ReleaseDC hNull, hDC
    
    ' Create the palette
    MakeIdentityPalette = CreatePalette(logpal)
End With
End Function

' Returns handle of a new palette identical to one passed as a parameter
Function DuplicatePalette(ByVal hPal As Long) As Long
    
    Dim f As Long, hDC As Long, cPal As Long
    hDC = GetDC(hNull)
    Dim logpal As LOGPALETTE256
    With logpal
        ' Get the size of the palette and dimension the array to it
        cPal = PalSize(hPal)
        .palVersion = &H300
        .palNumEntries = cPal
       
        ' Fill the logical palette array with palette colors
        f = GetPaletteEntries(hPal, 0, cPal, .palPalEntry(0))
        
        ReleaseDC hNull, hDC
        
        ' Create the palette
        DuplicatePalette = CreatePalette(logpal)
    End With
End Function

' Rotates an array of palette colors initialized by the CPalette class
Sub RotatePaletteArray(aColors() As OLE_COLOR, ByVal ecdA As ECycleDirection)
    Dim i As Long, clrT As OLE_COLOR, iLo As Long, iHi As Long
    iLo = LBound(aColors): iHi = UBound(aColors)
    
    Select Case ecdA
    Case ecdCycleLeft
        ' Left to right
        clrT = aColors(iLo)
        For i = iLo To iHi - 1
            aColors(i) = aColors(i + 1)
        Next
        aColors(iHi) = clrT
    Case ecdCycleRight
        ' Right to left
        clrT = aColors(iHi)
        For i = iHi To iLo + 1 Step -1
            aColors(i) = aColors(i - 1)
        Next
        aColors(iLo) = clrT
    Case ecdCycleIn
        ' In to the middle
        iHi = iHi \ 2
        ' Right to left
        clrT = aColors(iHi)
        For i = iHi To iLo + 1 Step -1
            aColors(i) = aColors(i - 1)
        Next
        aColors(iLo) = clrT
        ' Reset bounds
        iLo = iHi + 1
        iHi = UBound(aColors)
        ' Left to right
        clrT = aColors(iLo)
        For i = iLo To iHi - 1
            aColors(i) = aColors(i + 1)
        Next
        aColors(iHi) = clrT
        
    Case ecdCycleOut
        ' Out from the middle
        iHi = iHi \ 2
        ' Left to right
        clrT = aColors(iLo)
        For i = iLo To iHi - 1
            aColors(i) = aColors(i + 1)
        Next
        aColors(iHi) = clrT
        ' Reset bounds
        iLo = iHi + 1
        iHi = UBound(aColors)
        ' Right to left
        clrT = aColors(iHi)
        For i = iHi To iLo + 1 Step -1
            aColors(i) = aColors(i - 1)
        Next
        aColors(iLo) = clrT
    End Select
End Sub

' Add more palette manipulation functions here--start with palette fade in
' and fade out

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".PalTool"
        Select Case e
        Case eeBasePalTool
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

