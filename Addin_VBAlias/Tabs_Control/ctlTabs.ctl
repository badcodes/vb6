VERSION 5.00
Begin VB.UserControl ctlTabs 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   MouseIcon       =   "ctlTabs.ctx":0000
   PropertyPages   =   "ctlTabs.ctx":0152
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
End
Attribute VB_Name = "ctlTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'----------------------------------------------------------
' Tabs Controller (TabStrip Replacement)
'----------------------------------------------------------
' Author : Nick Gisburne
' Email  : nick@gisburne.com
' Web    : www.gisburne.com / www.karaokebuilder.com
'----------------------------------------------------------
' Purpose:
' Replacement for much of the functionality of the
' standard TabStrip control
'----------------------------------------------------------
' Limitations:
' I didn't set out to create an all-singing all-dancing
' tabs control. This is a simple control which displays
' tabs in a particular way, and is flexible enough for
' many purposes. If you want more, extend it yourself!
'----------------------------------------------------------
' Using the TabStrip supplied by Microsoft, I soon realised
' I wasn't getting much value for the 1-Mb overhead needed
' by the common controls ActiveX.  I also wanted the look
' and feel you see here, which I couldn't find elsewhere.
' I looked at other replacements for TabStrip, but they
' used quite a lot of resources (images, text boxes, etc)
' which all put a drain on Windows. This control just uses
' a totally empty control and draws on it. It's as simple
' as I could make it.
'
' I also added a PropertyPage - I've not used them before
' but it was surprisingly straightforward. Very useful
' for administering the various tabs and their captions.
'
'----------------------------------------------------------
' This control will be used in my commercial software, so
' if I think it's good enough I hope you do too! Enjoy!
'----------------------------------------------------------


'----------------------------------------------------------
' API Functions
'----------------------------------------------------------
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Integer
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'----------------------------------------------------------
' API Constants and Types
'----------------------------------------------------------
Const BDR_RAISEDOUTER = &H1
Const BDR_SUNKENOUTER = &H2
Const BDR_RAISEDINNER = &H4
Const BDR_SUNKENINNER = &H8

Const BDR_OUTER = &H3
Const BDR_INNER = &HC
Const BDR_RAISED = &H5
Const BDR_SUNKEN = &HA

Const BF_LEFT = &H1
Const BF_TOP = &H2
Const BF_RIGHT = &H4
Const BF_BOTTOM = &H8
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

'----------------------------------------------------------
' Enums for control properties
'----------------------------------------------------------
Public Enum ctlTabs_TabStyles
    [Top Left] = 0
    [Top Right] = 1
    [Bottom Left] = 2
    [Bottom Right] = 3
    [Left Top] = 4
    [Right Top] = 5
    [Left Bottom] = 6
    [Right Bottom] = 7
End Enum

Public Enum ctlTabs_CaptionStyle
    [cTop Left] = 0
    [cTop Center] = 1
    [cTop Right] = 2
    [cMiddle Left] = 3
    [cMiddle Center] = 4
    [cMiddle Right] = 5
    [cBottom Left] = 6
    [cBottom Center] = 7
    [cBottom Right] = 8
End Enum

'----------------------------------------------------------
' Variables used to store the control's properties
'----------------------------------------------------------
Dim propTabWide As Long
Dim propTabHigh As Long
Dim propTabSelected As Integer
Dim propTabCount As Integer
Dim propStyle As Integer
Dim propCaptionStyle As Integer
Dim propFocusRect As Boolean
Dim propCaption() As String
Dim propTabFont As StdFont
Dim propTabFontActive As StdFont
Dim propTabColor As OLE_COLOR
Dim propTabColorActive As OLE_COLOR
Dim propTextColor As OLE_COLOR
Dim propTextColorActive As OLE_COLOR

'----------------------------------------------------------
' Variables used for other purposes
'----------------------------------------------------------
Dim hasFocus As Boolean
Dim ClickZone() As RECT

'----------------------------------------------------------
' Event Declaration
'----------------------------------------------------------
Public Event TabClick(OldTab As Integer, NewTab As Integer)


'----------------------------------------------------------
' Properties Code (Get/Let/Set functions)
'----------------------------------------------------------

'----------------------------------------------------------
' ScaleMode / hDC
' Not really necessary but I had use for them
'----------------------------------------------------------
Public Property Get ScaleMode() As Long
ScaleMode = vbTwips     'Force to twips (it's actually set to pixels)
End Property

Public Property Get hDC() As Long
hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
hWnd = UserControl.hWnd
End Property
'----------------------------------------------------------
' TextColor / TextColorActive
' Color of text on the Inactive/Active tabs
'----------------------------------------------------------
Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Returns/sets the text color of all inactive tabs"
TextColor = propTextColor
End Property

Public Property Let TextColor(ByVal newVal As OLE_COLOR)
propTextColor = newVal
PropertyChanged "TEXTCOLOR"
DrawTabs
End Property

Public Property Get TextColorActive() As OLE_COLOR
Attribute TextColorActive.VB_Description = "Returns/sets the text color of the active tab"
TextColorActive = propTextColorActive
End Property

Public Property Let TextColorActive(ByVal newVal As OLE_COLOR)
propTextColorActive = newVal
PropertyChanged "TEXTCOLORACTIVE"
DrawTabs
End Property


'----------------------------------------------------------
' TabColor / TabColorActive
' Background Color of the Inactive/Active tabs
'----------------------------------------------------------
Public Property Get TabColor() As OLE_COLOR
Attribute TabColor.VB_Description = "Returns/sets the background color of all inactive tabs"
TabColor = propTabColor
End Property

Public Property Let TabColor(ByVal newVal As OLE_COLOR)
propTabColor = newVal
PropertyChanged "TABCOLOR"
DrawTabs
End Property

Public Property Get TabColorActive() As OLE_COLOR
Attribute TabColorActive.VB_Description = "Returns/sets the background color of the active tab and the container area of the control"
TabColorActive = propTabColorActive
End Property

Public Property Let TabColorActive(ByVal newVal As OLE_COLOR)
propTabColorActive = newVal
PropertyChanged "TABCOLORACTIVE"
DrawTabs
End Property


'----------------------------------------------------------
' TabFont / TabFontActive
' Font attributes of the Inactive/Active tabs
'----------------------------------------------------------
Public Property Get TabFont() As StdFont
Attribute TabFont.VB_Description = "Returns/sets the font attributes of all inactive tabs"
Set TabFont = propTabFont
End Property

Public Property Set TabFont(ByVal newVal As StdFont)
AssignFont propTabFont, newVal
PropertyChanged "TABFONT"
DrawTabs
End Property


Public Property Get TabFontActive() As StdFont
Attribute TabFontActive.VB_Description = "Returns/sets the font attributes of the active tab"
Set TabFontActive = propTabFontActive
End Property

Public Property Set TabFontActive(ByVal newVal As StdFont)
AssignFont propTabFontActive, newVal
PropertyChanged "TABFONTACTIVE"
DrawTabs
End Property


'----------------------------------------------------------
' Caption
' Tab captions
'----------------------------------------------------------
Public Property Get Caption(ByVal Index As Integer) As String
Attribute Caption.VB_Description = "Returns/Sets captions for each tab"
Caption = propCaption(Index)
End Property

Public Property Let Caption(ByVal Index As Integer, ByVal newVal As String)
propCaption(Index) = newVal
PropertyChanged "TABCAPTION" & Index
DrawTabs
End Property


'----------------------------------------------------------
' Style
' Where the tabs are displayed on the control
'----------------------------------------------------------
Public Property Get Style() As ctlTabs_TabStyles
Attribute Style.VB_Description = "Returns/sets the position of the tabs on the control"
Style = propStyle
End Property

Public Property Let Style(ByVal newVal As ctlTabs_TabStyles)
If newVal >= 0 And newVal <= 7 Then
    propStyle = newVal
    PropertyChanged "TABSTYLE"
    DrawTabs
End If
End Property


'----------------------------------------------------------
' CaptionAlignment
' Where the captions are displayed on the tabs
'----------------------------------------------------------
Public Property Get CaptionAlignment() As ctlTabs_CaptionStyle
Attribute CaptionAlignment.VB_Description = "Returns/sets the position of captions within the tabs"
CaptionAlignment = propCaptionStyle
End Property

Public Property Let CaptionAlignment(ByVal newVal As ctlTabs_CaptionStyle)
If newVal >= 0 And newVal <= 8 Then
    propCaptionStyle = newVal
    PropertyChanged "CAPTIONSTYLE"
    DrawTabs
End If
End Property


'----------------------------------------------------------
' TabsWidth / TabsHeight
' Dimensions of each tab (in pixels)
'----------------------------------------------------------
Public Property Get TabsWidth() As Long
Attribute TabsWidth.VB_Description = "Returns/sets the width of each tab (in pixels)"
TabsWidth = propTabWide
End Property

Public Property Let TabsWidth(ByVal newVal As Long)
propTabWide = newVal
PropertyChanged "TABWIDE"
DrawTabs
End Property

Public Property Get TabsHeight() As Long
Attribute TabsHeight.VB_Description = "Returns/sets the height of each tab (in pixels)"
TabsHeight = propTabHigh
End Property

Public Property Let TabsHeight(ByVal newVal As Long)
propTabHigh = newVal
PropertyChanged "TABHIGH"
DrawTabs
End Property


'----------------------------------------------------------
' Tabs
' The number of tabs on the control
'----------------------------------------------------------
Public Property Get Tabs() As Integer
Attribute Tabs.VB_Description = "Returns/sets the number of tabs on the control"
Tabs = propTabCount
End Property

Public Property Let Tabs(ByVal newVal As Integer)
Dim oldVal As Integer, tabChanged As Boolean
If newVal > 0 Then
    propTabCount = newVal
    ReDim Preserve propCaption(1 To propTabCount)
'Reducing the number of tabs can also change the selected tab
    If propTabCount < propTabSelected Then
        oldVal = propTabSelected
        propTabSelected = propTabCount
        tabChanged = True
    End If
    PropertyChanged "TABCOUNT"
    DrawTabs
'Do this here because we want to raise the event
'AFTER the tabs have been drawn
    If tabChanged Then
        PropertyChanged "TABSELECTED"
        RaiseEvent TabClick(oldVal, propTabSelected)
    End If
End If
End Property


'----------------------------------------------------------
' SelectedTab
' Currently selected (active) tab
'----------------------------------------------------------
Public Property Get SelectedTab() As Integer
Attribute SelectedTab.VB_Description = "Returns/sets the currently selected tab"
SelectedTab = propTabSelected
End Property

Public Property Let SelectedTab(ByVal newVal As Integer)
Dim oldVal As Integer

oldVal = propTabSelected
propTabSelected = newVal

If propTabSelected < 1 Then         'Range checks
    propTabSelected = 1
ElseIf propTabSelected > propTabCount Then
    propTabSelected = propTabCount
End If

PropertyChanged "TABSELECTED"
DrawTabs
'Do this here because we want to raise the event
'AFTER the tabs have been drawn
RaiseEvent TabClick(oldVal, newVal)
End Property


'----------------------------------------------------------
' UseFocusRect
' Indicates the focus should be shown
'----------------------------------------------------------
Public Property Get UseFocusRect() As Boolean
Attribute UseFocusRect.VB_Description = "Returns/sets whether a focus rectangle will be displayed when a tab has the focus"
UseFocusRect = propFocusRect
End Property

Public Property Let UseFocusRect(ByVal newVal As Boolean)
propFocusRect = newVal
PropertyChanged "FOCUSRECT"
DrawTabs
End Property


'----------------------------------------------------------
' UserControl Functions
'----------------------------------------------------------

'----------------------------------------------------------
' Changes background color to that of your color scheme
'----------------------------------------------------------
Private Sub UserControl_AmbientChanged(PropertyName As String)
If PropertyName = "BackColor" Then DrawTabs
End Sub

'----------------------------------------------------------
' Arrow keys can be used to move between tabs
'----------------------------------------------------------
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyLeft, vbKeyUp:    SelectedTab = SelectedTab - 1
Case vbKeyRight, vbKeyDown: SelectedTab = SelectedTab + 1
End Select
End Sub

'----------------------------------------------------------
' Handles the focus rectangle
'----------------------------------------------------------
Private Sub UserControl_GotFocus()
hasFocus = True
If propFocusRect Then DrawTabs
End Sub

Private Sub UserControl_LostFocus()
hasFocus = False
If propFocusRect Then DrawTabs
End Sub

'----------------------------------------------------------
' If you click on one of the tabs, you select that tab
'----------------------------------------------------------
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t1 As Integer
For t1 = 1 To propTabCount
    If X >= ClickZone(t1).Left And X <= ClickZone(t1).Right _
            And Y >= ClickZone(t1).Top And Y <= ClickZone(t1).Bottom Then
        SelectedTab = t1
        Exit For
    End If
Next t1
End Sub

'----------------------------------------------------------
' If your mouse moves over one of the tabs
' the MousePointer becomes a hand pointer
'----------------------------------------------------------
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t1 As Integer, tPointer As Integer

tPointer = vbDefault
For t1 = 1 To propTabCount
    If X >= ClickZone(t1).Left And X <= ClickZone(t1).Right _
            And Y >= ClickZone(t1).Top And Y <= ClickZone(t1).Bottom Then
        tPointer = vbCustom
        Exit For
    End If
Next t1
UserControl.MousePointer = tPointer
End Sub

Private Sub UserControl_Resize()
DrawTabs
End Sub

'----------------------------------------------------------
' Initialize - Default values for all properties
'----------------------------------------------------------
Private Sub UserControl_Initialize()
Dim t1 As Integer
propTabWide = 70: propTabHigh = 20
propTabCount = 2
propTabSelected = 1
propStyle = [Top Left]
propCaptionStyle = [cMiddle Center]
propFocusRect = False
ReDim propCaption(1 To propTabCount)
For t1 = 1 To propTabCount
    propCaption(t1) = "Tab " & t1
Next t1
Set propTabFont = New StdFont
AssignFont propTabFont, UserControl.Font
Set propTabFontActive = New StdFont
AssignFont propTabFontActive, UserControl.Font
propTabFontActive.Bold = Not propTabFontActive.Bold
propTabColor = vbButtonShadow
propTabColorActive = vbButtonFace
propTextColor = vb3DHighlight
propTextColorActive = vbButtonText
End Sub

'----------------------------------------------------------
' ReadProperties
'----------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim t1 As Integer
With PropBag
    propTabWide = .ReadProperty("TABWIDE", 70)
    propTabHigh = .ReadProperty("TABHIGH", 20)
    propTabCount = .ReadProperty("TABCOUNT", 2)
    propTabSelected = .ReadProperty("TABSELECTED", 1)
    propStyle = .ReadProperty("TABSTYLE", [Top Left])
    propCaptionStyle = .ReadProperty("CAPTIONSTYLE", [cMiddle Center])
    propFocusRect = .ReadProperty("FOCUSRECT", False)
    ReDim Preserve propCaption(1 To propTabCount)
    For t1 = 1 To propTabCount
        propCaption(t1) = .ReadProperty("TABCAPTION" & t1, "")
    Next t1
    Set propTabFont = .ReadProperty("TABFONT", UserControl.Font)
    Set propTabFontActive = .ReadProperty("TABFONTACTIVE", UserControl.Font)
    propTabColor = .ReadProperty("TABCOLOR", vbButtonShadow)
    propTabColorActive = .ReadProperty("TABCOLORACTIVE", vbButtonFace)
    propTextColor = .ReadProperty("TEXTCOLOR", vb3DHighlight)
    propTextColorActive = .ReadProperty("TEXTCOLORACTIVE", vbButtonText)
End With
DrawTabs
End Sub

'----------------------------------------------------------
' WriteProperties
'----------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim t1 As Integer
With PropBag
    .WriteProperty "TABWIDE", propTabWide
    .WriteProperty "TABHIGH", propTabHigh
    .WriteProperty "TABCOUNT", propTabCount
    .WriteProperty "TABSELECTED", propTabSelected
    .WriteProperty "TABSTYLE", propStyle
    .WriteProperty "CAPTIONSTYLE", propCaptionStyle
    .WriteProperty "FOCUSRECT", propFocusRect
    For t1 = 1 To propTabCount
        .WriteProperty "TABCAPTION" & t1, propCaption(t1), ""
    Next t1
    .WriteProperty "TABFONT", propTabFont
    .WriteProperty "TABFONTACTIVE", propTabFontActive
    .WriteProperty "TABCOLOR", propTabColor
    .WriteProperty "TABCOLORACTIVE", propTabColorActive
    .WriteProperty "TEXTCOLOR", propTextColor
    .WriteProperty "TEXTCOLORACTIVE", propTextColorActive
End With
End Sub


'----------------------------------------------------------
' DrawTabs
' The main routine which draws everything on the control
'----------------------------------------------------------
Private Sub DrawTabs()
Dim tPoint As POINTAPI, trect As RECT, tRectIn As RECT
Dim tBorderButton As Integer
Dim t1 As Integer, tCaption As String
Dim tCol As Long, hBrushInactive As Long, hBrushActive As Long
Dim tTextCol As Long, tTextActive As Long

Dim tCtlHeight As Long, tCtlWidth As Long   'Control's height/width in pixels

'Color translation is required - eg to convert 'Button Text'
'to the RGB value used by your Windows theme
OleTranslateColor propTextColor, 0, tTextCol
OleTranslateColor propTextColorActive, 0, tTextActive

'Brushes used to draw filled rectangles
OleTranslateColor propTabColor, 0, tCol:        hBrushInactive = CreateSolidBrush(tCol)
OleTranslateColor propTabColorActive, 0, tCol:  hBrushActive = CreateSolidBrush(tCol)

With UserControl
    tCtlWidth = .Width / Screen.TwipsPerPixelX
    tCtlHeight = .Height / Screen.TwipsPerPixelY

    .Backcolor = UserControl.Ambient.Backcolor      'Non-tabs area matches your Windows theme
    .Cls

'----------------------------------------------------------
' Draw rectangle for the main container area
'----------------------------------------------------------
    trect.Left = 0
    trect.Right = tCtlWidth
    trect.Top = 0
    trect.Bottom = tCtlHeight

'tBorderButton - these styles indicate rectangles with one side missing
    Select Case propStyle
    Case [Top Left], [Top Right]
        trect.Top = propTabHigh
        tBorderButton = BF_RECT - BF_BOTTOM

    Case [Bottom Left], [Bottom Right]
        trect.Bottom = tCtlHeight - propTabHigh
        tBorderButton = BF_RECT - BF_TOP

    Case [Left Top], [Left Bottom]
        trect.Left = propTabWide
        tBorderButton = BF_RECT - BF_RIGHT

    Case [Right Top], [Right Bottom]
        trect.Right = tCtlWidth - propTabWide
        tBorderButton = BF_RECT - BF_LEFT
    End Select

    FillRect .hDC, trect, hBrushActive              'Filled rectangle
    DrawEdge .hDC, trect, BDR_RAISEDINNER, BF_RECT  'Surround with 3D edge

'----------------------------------------------------------
' Draw each tab onto the remainder of the control
'----------------------------------------------------------
    ReDim ClickZone(1 To propTabCount)  'Stored the coords of each tab

    For t1 = 1 To propTabCount
        Select Case propStyle
        Case [Top Left]
            trect.Left = (t1 - 1) * propTabWide
            trect.Top = 0

        Case [Top Right]
            trect.Left = tCtlWidth - propTabCount * propTabWide + (t1 - 1) * propTabWide - 1
            trect.Top = 0

        Case [Bottom Left]
            trect.Left = (t1 - 1) * propTabWide
            trect.Top = tCtlHeight - propTabHigh - 1

        Case [Bottom Right]
            trect.Left = tCtlWidth - propTabCount * propTabWide + (t1 - 1) * propTabWide - 1
            trect.Top = tCtlHeight - propTabHigh - 1

        Case [Left Top]
            trect.Left = 0
            trect.Top = (t1 - 1) * propTabHigh

        Case [Left Bottom]
            trect.Left = 0
            trect.Top = tCtlHeight - propTabCount * propTabHigh + (t1 - 1) * propTabHigh - 1

        Case [Right Top]
            trect.Left = tCtlWidth - propTabWide - 1
            trect.Top = (t1 - 1) * propTabHigh

        Case [Right Bottom]
            trect.Left = tCtlWidth - propTabWide - 1
            trect.Top = tCtlHeight - propTabCount * propTabHigh + (t1 - 1) * propTabHigh - 1

        End Select
        trect.Right = trect.Left + propTabWide + 1
        trect.Bottom = trect.Top + propTabHigh + 1

'----------------------------------------------------------
' Draw the selected (active) tab
'----------------------------------------------------------
        If t1 = propTabSelected Then
            FillRect .hDC, trect, hBrushActive
            DrawEdge .hDC, trect, BDR_RAISEDINNER, tBorderButton
            ClickZone(t1) = trect
            If propFocusRect And hasFocus Then
                tRectIn.Left = trect.Left + 2       'Focus rectangle
                tRectIn.Right = trect.Right - 2
                tRectIn.Top = trect.Top + 2
                tRectIn.Bottom = trect.Bottom - 2
                DrawFocusRect .hDC, tRectIn
            End If
            AssignFont .Font, propTabFontActive
            .ForeColor = propTextColorActive
'----------------------------------------------------------
' Draw the inactive tabs
'----------------------------------------------------------
        Else
            tRectIn.Left = trect.Left + 1           'Rectangle for inactive tabs is smaller
            tRectIn.Right = trect.Right - 1
            tRectIn.Top = trect.Top + 1
            tRectIn.Bottom = trect.Bottom - 1
            FillRect .hDC, tRectIn, hBrushInactive
            ClickZone(t1) = tRectIn
            AssignFont .Font, propTabFont
            .ForeColor = propTextColor
        End If

'----------------------------------------------------------
' Draw the caption on each tab
'----------------------------------------------------------
        tCaption = propCaption(t1)
        Select Case propCaptionStyle    'X coord
        Case 0, 3, 6:   .CurrentX = trect.Left + 3
        Case 1, 4, 7:   .CurrentX = (trect.Right - trect.Left) / 2 + trect.Left - .TextWidth(tCaption) / 2 - 1
        Case 2, 5, 8:   .CurrentX = trect.Right - .TextWidth(tCaption) - 4
        End Select
        Select Case propCaptionStyle    'Y coord
        Case 0, 1, 2:   .CurrentY = trect.Top + 2
        Case 3, 4, 5:   .CurrentY = (trect.Bottom - trect.Top) / 2 + trect.Top - .TextHeight(tCaption) / 2 - 1
        Case 6, 7, 8:   .CurrentY = (trect.Bottom) - .TextHeight(tCaption) - 2
        End Select

        UserControl.Print tCaption

    Next t1
End With

DeleteObject hBrushInactive
DeleteObject hBrushActive
End Sub


'----------------------------------------------------------
' AssignFont
' Copy the attributes of one font object to another
'----------------------------------------------------------
Private Sub AssignFont(ByRef tFont1 As StdFont, ByVal tfont2 As StdFont)
With tFont1
    .Bold = tfont2.Bold
    .Charset = tfont2.Charset
    .Italic = tfont2.Italic
    .Name = tfont2.Name
    .Size = tfont2.Size
    .Strikethrough = tfont2.Strikethrough
    .Underline = tfont2.Underline
    .Weight = tfont2.Weight
End With
End Sub
