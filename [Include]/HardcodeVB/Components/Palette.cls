VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Public Enum ECycleDirection2
    ecdCycleLeft
    ecdCycleRight
End Enum

Private hPal As Long, hPalOld As Long
Private hWnd As Long, hDC As Long, cPal As Long, cPalReal As Long
Private iFrom As Long, iTo As Long
Private ape() As PALETTEENTRY
Private emr As EMsgResponse

Public Enum EErrorPalette
    eeBasePalette = 13130   ' CPalette
    eeInvalidPalette        ' Invalid palette
    eePaletteNotInit        ' Palette not initialized with Create
    eeBitmapNoPalette       ' Bitmap has no palette
    eeInvalidExclusion      ' Invalid first or last index
    eeCantResizeArray       ' Input array must be resizable
End Enum

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".CPalette"
        Select Case e
        Case eeBasePalette
            BugAssert True
        Case eeInvalidPalette
            sText = "Invalid palette"
        Case eePaletteNotInit
            sText = "Palette not initialized with Create"
        Case eeBitmapNoPalette
            sText = "Bitmap has no palette"
        Case eeInvalidExclusion
            sText = "Invalid first or last index"
        Case eeCantResizeArray
            sText = "Input array must be resizable"
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

Public Function Create(ByVal hPalA As Long, ByVal hWndA As Long, _
                       aColors() As OLE_COLOR, _
                       Optional FirstIndex As Long = 0, _
                       Optional LastIndex As Long = -1) As Long
                       
    ' Must be a new palette handle and same old window
    If hPalA = hNull Or hPalA = hInvalid Then ErrRaise eeBitmapNoPalette
    If hWndA = hNull Then ApiRaise ERROR_INVALID_HANDLE
    Destroy
    ' Initialize members
    hPalOld = hPalA
    hPal = hNull
    iFrom = 0
    hWnd = hWndA
    hDC = GetDC(hWnd)
    ' Get the size
    cPalReal = MPalTool.PalSize(hPalOld)
    ' Dimension an array for the real palette
    ReDim ape(0 To cPalReal - 1) As PALETTEENTRY
    ' Create a dumplicate palette
    hPal = MPalTool.DuplicatePalette(hPalOld)
    If hPal = hNull Then ErrRaise eeInvalidPalette
    ' Adjust the exclusions and calculate the excluded length
    iFrom = FirstIndex
    If LastIndex = -1 Then
        iTo = cPalReal - iFrom - 1
    Else
        iTo = LastIndex
    End If
    If iFrom >= iTo Or iTo - iFrom + 1 > cPalReal Then
        ErrRaise eeInvalidExclusion
    End If
    cPal = iTo - iFrom + 1
    ' Get the palette entries, mark them reserved, and save result
    Dim i As Long, c As Long
    c = GetPaletteEntries(hPal, 0, cPalReal, ape(0))
    BugAssert c = cPalReal
    For i = iFrom To iTo
        ape(i).peFlags = PC_RESERVED
    Next
    c = SetPaletteEntries(hPal, 0, cPalReal, ape(0))
    BugAssert c = cPalReal
    If c = 0 Then BugMessage "Fail SetPaletteEntries: " & Err.LastDllError
    ' Realize this new palette
    Realize
    ' Initialize the user's work palette
    On Error GoTo CreateFail
    ReDim aColors(0 To cPal - 1) As OLE_COLOR
    CopyMemory aColors(0), ape(iFrom), cPal * 4
    ' Subclass window to handle palette messages
    AttachMessage Me, hWnd, WM_PALETTECHANGED
    AttachMessage Me, hWnd, WM_QUERYNEWPALETTE
    ' Return real color count
    Create = cPalReal
    FirstIndex = iFrom
    LastIndex = iTo
    Exit Function
CreateFail:
    If Err <> 0 And Err.Number = eeArrayLocked Then
        ErrRaise eeCantResizeArray
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Function

Public Function Destroy()
    ' Detach messages, restore old palette, free the user's work array
    DetachMessage Me, hWnd, WM_QUERYNEWPALETTE
    DetachMessage Me, hWnd, WM_PALETTECHANGED
    SelectPalette hDC, hPalOld, APIFALSE
    Call RealizePalette(hDC)
    hDC = ReleaseDC(hWnd, hDC)
    Erase ape
    hPal = 0
    hDC = 0
    hWnd = 0
    cPal = 0
    cPalReal = 0
End Function

Private Sub Class_Initialize()
    iTo = -1
    iFrom = -1
End Sub

Private Sub Class_Terminate()
    ' Always destroy
    Destroy
End Sub

' Our work palette
Public Property Get Handle() As Long
    Handle = hPal
End Property

' User's original palette
Public Property Get SourceHandle() As Long
    SourceHandle = hPalOld
End Property

Public Property Get RGBColor(ByVal i As Long) As OLE_COLOR
    RGBColor = PaletteColorFromEntry(ape(i))
End Property
Public Property Let RGBColor(ByVal i As Long, ByVal clr As OLE_COLOR)
    If hPal = 0 Then ErrRaise eePaletteNotInit
    Call PaletteColorToEntry(ape(i), clr)
End Property

' Size of our work palette
Property Get Size() As Long
    Size = cPal
End Property

' Size of user's original palette
Property Get SourceSize() As Long
    SourceSize = cPalReal
End Property

Property Get PaletteColor(ByVal i As Long) As OLE_COLOR
    PaletteColor = &H2000000 Or PaletteColorFromEntry(ape(i))
End Property

' Caller passes in a modified array to be written to the palette
Sub ModifyPalette(aColors() As OLE_COLOR)
    If hPal = 0 Then ErrRaise eePaletteNotInit
    CopyMemory ape(iFrom), aColors(0), cPal * 4
    Animate
End Sub

' Animate the modified palette
Sub Animate()
    If hPal = 0 Then ErrRaise eePaletteNotInit
    Dim f As Long
    f = AnimatePalette(hPal, iFrom, cPal, ape(iFrom))
    If f = 0 Then BugMessage "Fail AnimatePalette: " & Err.LastDllError
End Sub

' Select our our palette and realize it into the system
Private Sub Realize(Optional Background As Boolean = False)
    If hPal = 0 Then ErrRaise eePaletteNotInit
    Dim h As Long
    h = SelectPalette(hDC, hPal, -Background)
    If h Then BugMessage "Fail SelectPalette: " & Err.LastDllError
    Dim c As Long
    c = RealizePalette(hDC)
    If c = GDI_ERROR Then BugMessage "Fail RealizePalette: " & Err.LastDllError
End Sub

' Implement ISubclass

Private Property Let ISubclass_MsgResponse(ByVal emrA As EMsgResponse)
    emr = emrA
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emr
End Property

Private Function ISubclass_WindowProc(ByVal hWndA As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
    ' Handle message
    Select Case iMsg
    Case WM_PALETTECHANGED
        ' Background
        If wParam <> hWndA Then Realize False
        emr = emrPostProcess
    Case WM_QUERYNEWPALETTE
        ' Foreground
        Realize True
        ISubclass_WindowProc = APITRUE
        emr = emrConsume
    End Select
End Function

' ---- Private Helpers ----
'
Private Function PaletteColorFromEntry(pe As PALETTEENTRY) As OLE_COLOR
    ' Copy color bytes, ignore flag byte
    CopyMemory PaletteColorFromEntry, pe, 3
End Function

Private Sub PaletteColorToEntry(pe As PALETTEENTRY, ByVal clr As OLE_COLOR)
    ' Copy color bytes, ignore flag byte
    CopyMemory pe, clr, 3
End Sub




