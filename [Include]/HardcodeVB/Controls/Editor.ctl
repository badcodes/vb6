VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl XEditor 
   Alignable       =   -1  'True
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   PaletteMode     =   4  'None
   ScaleHeight     =   1830
   ScaleWidth      =   2190
   ToolboxBitmap   =   "Editor.ctx":0000
   Begin RichTextLib.RichTextBox txt 
      Height          =   1572
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1932
      _ExtentX        =   3413
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Editor.ctx":00FA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "XEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RepId" ,"4FB51841-CEAF-11CF-A15E-00AA00A74D48-005c"
Option Explicit

Public Enum EErrorEditor
    eeBaseEditor = 13720    ' XEditor
End Enum

Public Enum ELoadSave
    elsDefault = -1
    elsrtf
    elstext
End Enum

Public Enum EWordWrap
    NoWordWrap = 65535
End Enum

Public Enum ESearchEvent
    eseFindWhat
    eseReplaceWith
    eseCase
    eseWholeWord
    eseDirection
End Enum

' This isn't passed through from RichTextBox
Public Enum EAlignment
    rtfLeft
    rtfRight
    rtfCenter
End Enum

Public Enum ESearchDir
    esdAll
    esdDown
    esdUp
End Enum
Private esdDir As ESearchDir

' Private variables for properties
Private cCharPerTab As Integer
Private sFilter As String, iFilter As Integer
Private sFilePath As String
Private fTextMode As Boolean
Private fSaveWordWrap As Boolean
Private fEnableTab As Boolean
Private ordAppearance As AppearanceConstants
Private ordScrollBars As ScrollBarsConstants
Private fSearchOptionCase As Boolean
Private fSearchOptionWord As Boolean  ' Left as an exercise
Private cSearchActive As Integer
Private fOverWrite As Boolean
Private nFilters As New Collection
Private xMin As Single, yMin As Single
Private clrFore As OLE_COLOR
Private vecTab As New CVectorBool
Private cFindWhatMax As Long
Private cReplaceWithMax As Long
Private nFindWhat As New Collection
Private nReplaceWith As New Collection
Private fontDefault As Font
' Dialogs at module level so they can be destroyed
Private finddlg As New FSearch

'Event Declarations

' RichTextBox events
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses a mouse button."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user presses and releases a mouse button."
Event Change()
Attribute Change.VB_Description = "Indicates that the contents of a control have changed."
Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "OLECompleteDrag event"
Event OLEDragDrop(data As RichTextLib.DataObject, Effect As Long, _
                  Button As Integer, Shift As Integer, _
                  x As Single, y As Single)
Event OLEDragOver(data As RichTextLib.DataObject, Effect As Long, _
                  Button As Integer, Shift As Integer, _
                  x As Single, y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(data As RichTextLib.DataObject, DataFormat As Integer)
Event OLEStartDrag(data As RichTextLib.DataObject, AllowedEffects As Long)
Event SearchChange(Kind As ESearchEvent)
Event SelChange()

' New event that reports status
Event StatusChange(LineCur As Long, LineCount As Long, _
                   ColumnCur As Long, ColumnCount As Long, _
                   CharacterCur As Long, CharacterCount As Long, _
                   DirtyBit As Boolean)


Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".Utility"
        Select Case e
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.EXEName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
     BugLocalMessage "XEditor UserControl_AmbientChanged: " & PropertyName
End Sub

Private Sub UserControl_EnterFocus()
    BugLocalMessage "XEditor UserControl_EnterFocus"
End Sub

Private Sub UserControl_ExitFocus()
    BugLocalMessage "XEditor UserControl_ExitFocus"
End Sub

Private Sub UserControl_Initialize()
    BugLocalMessage "XEditor UserControl_Initialize"
    xMin = Width / 2
    yMin = Height / 2
    SearchOptionDirection = esdAll
End Sub

Private Sub UserControl_GotFocus()
    BugLocalMessage "XEditor UserControl_GotFocus"
End Sub

Private Sub UserControl_LostFocus()
    BugLocalMessage "XEditor UserControl_LostFocus"
End Sub

Private Sub txt_GotFocus()
    BugLocalMessage "XEditor txt_GotFocus"
    If fEnableTab And Ambient.UserMode Then
        ' Ignore errors for controls without the TabStop property
        On Error Resume Next
        Dim i As Long, vControl As Variant, f As Boolean
        ' Clear tab vector
        i = 1
        Set vecTab = Nothing
        ' Stop changing focus when pressing TAB
        For Each vControl In Parent.Controls
            f = vControl.TabStop
            vecTab(i) = f
            vControl.TabStop = False
            i = i + 1
        Next
    End If
End Sub

Private Sub txt_LostFocus()
    BugLocalMessage "XEditor txt_LostFocus"
    If fEnableTab And Ambient.UserMode Then
        ' Ignore errors for controls without the TabStop property
        On Error Resume Next
        Dim i As Long, vControl As Variant
        ' Restore TabStops from vector
        i = 1
        For Each vControl In Parent.Controls
            vControl.TabStop = vecTab(i)
            i = i + 1
        Next
    End If
End Sub

Private Sub UserControl_Terminate()
    BugLocalMessage "XEditor UserControl_Terminate"
    ' If a find or replace dialog is active, terminate it
    Set finddlg = Nothing
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    BugLocalMessage "XEditor UserControl_InitProperties"
    fTextMode = True
    ordAppearance = rtfThreeD
'    ordScrollBars = rtfBoth
#If afDebug Then
    ChDrive App.Path
    ChDir App.Path
#End If
    cCharPerTab = 8
    nFilters.Add "Text files (*.txt): *.txt"
    nFilters.Add "Rich text files (*.rtf): *.rtf"
    nFilters.Add "All files (*.*): *.*"
    Text = sEmpty
    Extender.Name = UniqueControlName("edit", Extender)
    UserControl_Load
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With txt
    BugLocalMessage "XEditor UserControl_ReadProperties"
    ordAppearance = PropBag.ReadProperty("Appearance", rtfThreeD)
    .AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", .AutoVerbMenu)
    .BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    .BorderStyle = PropBag.ReadProperty("BorderStyle", rtfFixedSingle)
    .BulletIndent = PropBag.ReadProperty("BulletIndent", 0)
    ' Enabled provided by extender
    sFilePath = PropBag.ReadProperty("FileName", sEmpty)
    Dim s As String, st As String
    s = PropBag.ReadProperty("FileOpenFilter", sEmpty)
    Set .Font = PropBag.ReadProperty("Font", Ambient.Font)
    TextColor = PropBag.ReadProperty("TextColor", vbWindowText)
    ' Height in extender
    ' HelpContextID on extender
    .HideSelection = PropBag.ReadProperty("HideSelection", False)
    ' hWnd run time only
    ' Index on extender
    ' Left in Extender
    ' Line, Lines run time only
    ' LinePosition, LineLength run time only
    ' LineText run time only
    .Locked = PropBag.ReadProperty("Locked", False)
    Set .MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    .MousePointer = PropBag.ReadProperty("MousePointer", rtfDefault)
    .OLEDragMode = PropBag.ReadProperty("OLEDragMode", rtfOLEDragAutomatic)
    .OLEDropMode = PropBag.ReadProperty("OLEDropMode", rtfOLEDropAutomatic)
    fOverWrite = PropBag.ReadProperty("OverWrite", False)
    ' Percent run time only
    .RightMargin = PropBag.ReadProperty("RightMargin", 0)
    ordScrollBars = PropBag.ReadProperty("ScrollBars", rtfBoth)
    ScrollBars = ordScrollBars
    ' SearchOptionDirection, SearchOptionCase, SearchOptionWord run time only
    ' SelAlignment run time only
    ' SelBold run time only
    ' SelBullet run time only
    ' SelCharOffset run time only
    ' SelColor run time only
    ' SelFontName run time only
    ' SelFontSize run time only
    ' SelHangingIndent run time only
    ' SelIndent run time only
    ' SelItalic run time only
    ' SelLength run time only
    ' SelProtected run time only
    ' SelRightIndent run time only
    ' SelRTF run time only
    ' SelStart run time only
    ' SelStrikeThru run time only
    ' SelTabCount run time only
    ' .SelTabs(sElement) = PropBag.ReadProperty("SelTabs" & Index, 0)
    ' SelTabs run time only
    ' SelText run time only
    ' SelUnderline run time only
    .Text = PropBag.ReadProperty("Text", sEmpty)
    ' TextRTF run time only
    TextMode = PropBag.ReadProperty("TextMode", True)
    UserControl_Load
End With
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    BugLocalMessage "XEditor UserControl_WriteProperties"
With txt
    PropBag.WriteProperty "Appearance", ordAppearance, rtfThreeD
    PropBag.WriteProperty "AutoVerbMenu", .AutoVerbMenu, False
    PropBag.WriteProperty "BackColor", .BackColor, vbWindowBackground
    PropBag.WriteProperty "BorderStyle", .BorderStyle, rtfFixedSingle
    PropBag.WriteProperty "BulletIndent", .BulletIndent, 0
    ' Character, Characters run time only
    ' Column, Columns run time only
    ' Container controlled by extender
    ' DirtyBit run time only
    ' Enabled handled by extender
    PropBag.WriteProperty "FileName", sFilePath, sEmpty
    Dim s As String, v As Variant
    'For Each v In nFilters
    '    s = s & v & sCrLf
    'Next
    's = Left$(s, Len(s) - 2)
    PropBag.WriteProperty "FileOpenFilter", s, sEmpty
    ' FindWhat, ReplaceWith run time only
    ' FindWhatList, ReplaceWithList run time only
    PropBag.WriteProperty "Font", .Font, Ambient.Font
    PropBag.WriteProperty "TextColor", TextColor, vbWindowText
    ' Height in extender
    ' hWnd run time only
    PropBag.WriteProperty "HideSelection", .HideSelection, False
    PropBag.WriteProperty "Locked", .Locked, False
    PropBag.WriteProperty "MouseIcon", .MouseIcon, Nothing
    PropBag.WriteProperty "MousePointer", .MousePointer, rtfDefault
    PropBag.WriteProperty "OLEDragMode", .OLEDragMode, rtfOLEDragAutomatic
    PropBag.WriteProperty "OLEDropMode", .OLEDropMode, rtfOLEDropAutomatic
    PropBag.WriteProperty "OverWrite", fOverWrite, False
    PropBag.WriteProperty "RightMargin", .RightMargin, 0
    PropBag.WriteProperty "ScrollBars", ordScrollBars, rtfBoth
    ' SelAlignment run time only
    ' SelBold run time only
    ' SelBullet run time only
    ' SelCharOffset run time only
    ' SelColor run time only
    ' SelFontName run time only
    ' SelFontSize run time only
    ' SelHangingIndent run time only
    ' SelIndent run time only
    ' SelItalic run time only
    ' SelLength run time only
    ' SelProtected run time only
    ' SelRightIndent run time only
    ' SelRTF run time only
    ' SelStart run time only
    ' SelStrikeThru run time only
    ' SelTabCount run time only
    ' SelTabs" & Index, .SelTabs(sElement), 0
    ' SelText run time only
    ' SelUnderline run time only
    PropBag.WriteProperty "Text", .Text, sEmpty
    PropBag.WriteProperty "TextMode", TextMode, True
End With
End Sub

Private Sub UserControl_Resize()
    BugLocalMessage "XEditor UserControl_Resize"
    If Width < xMin Then Width = xMin
    If Height < yMin Then Height = yMin
    ' Adjust internal RichTextBox to be the size of the UserControl
    txt.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_Load()
    DirtyBit = False
    InitFilters
    If fOverWrite Then
        ' Make overwrite state match fOverWrite variable
        SendMessage txt.hWnd, WM_KEYDOWN, ByVal VK_INSERT, ByVal &H510001
        SendMessage txt.hWnd, WM_KEYUP, ByVal VK_INSERT, ByVal &HC0510001
    End If
    Set fontDefault = Font
End Sub

'' New methods

Sub FileNew()
    Text = sEmpty
    sFilePath = sEmpty
    DirtyBit = False
End Sub

Function FileOpen() As Boolean
    Dim f As Boolean, sFile As String, fReadOnly As Boolean
    f = VBGetOpenFileName( _
            FileName:=sFile, _
            ReadOnly:=fReadOnly, _
            filter:=FilterString, _
            DefaultExt:=IIf(TextMode, ".txt", ".rtf"), _
            FilterIndex:=IIf(TextMode, 1, 2), _
            Owner:=hWnd)
    If f And sFile <> sEmpty Then
        TextMode = Not IsRTF(sFile)
        LoadFile sFile
        If fReadOnly Then Locked = True
        FileOpen = True
    End If
End Function

Property Get FileOpenFilter(Optional ByVal i As Integer) As String
    If i = 0 Then i = 1
    If i > nFilters.Count Or i < 1 Then Exit Property
    FileOpenFilter = nFilters(i)
End Property

Property Let FileOpenFilter(Optional ByVal i As Integer, _
                            sFilterA As String)
    If i > nFilters.Count Or i < 1 Then
        nFilters.Add sFilterA
    Else
        nFilters.Add sFilterA, , i
    End If
    PropertyChanged "FileOpenFilter"
End Property

Sub FileSave()
    If FileName = sEmpty Then
        FileSaveAs
    Else
        SaveFile FilePath
    End If
    DirtyBit = False
End Sub

Function FileSaveAs() As Boolean
    Dim f As Boolean, sFile As String
    ' Default includes OverWritePrompt so you can confirm
    f = VBGetSaveFileName( _
            FileName:=sFile, _
            filter:=FilterString, _
            DefaultExt:=IIf(TextMode, ".txt", ".rtf"), _
            FilterIndex:=IIf(TextMode, 1, 2), _
            Owner:=hWnd)
    If f And sFile <> sEmpty Then
        ' Be sure you're right, then go ahead
        If ExistFile(sFile) Then Kill sFile
        SaveFile sFile
        FileSaveAs = True
    End If
End Function

Sub FilePrint()
    Dim hDC As Long, epr As EPrintRange
    If VBPrintDlg(hDC, DisablePageNumbers:=True, _
                       DisableSelection:=SelLength = 0, _
                       PrintRange:=epr, _
                       ShowPrintToFile:=False, _
                       Owner:=hWnd) Then
        ' We don't handle request to print specific pages
        If epr = eprAll Then
            ' Print all regardless of selection
            SelPrint hDC, True
        Else
            SelPrint hDC
        End If
    End If
End Sub

Sub FilePageSetup()
    If VBPageSetupDlg(Owner:=hWnd) Then
        
    End If
End Sub

#If 1 Then
' OpenColor with my VBChooseColor function
Function OptionColor(Optional ByVal clr As Long = vbBlack) As Long
    ' Make sure it's an RGB color
    clr = TranslateColor(clr)
    ' Choose a solid color
    Call VBChooseColor(Color:=clr, AnyColor:=False, Owner:=hWnd)
    ' Return color, whether successful or not
    OptionColor = clr
End Function
#ElseIf 0 Then
' OpenColor as it would work with Dialog Automation Objects
Function OptionColor(Optional ByVal clr As Long = vbBlack) As Long
Dim choose As New ChooseColor
With choose
    ' Make sure it's an RGB color
    .Color = TranslateColor(clr)
    .hWnd = hWnd
    ' No property to specify solid colors
    ' Return color, whether successful or not
    If .Show Then
        OptionColor = choose.Color
    Else
        OptionColor = clr
    End If
End With
End Function
#ElseIf 0 Then
' OpenColor as it would work with the CommonDialog control
Function OptionColor(Optional ByVal clr As Long = vbBlack) As Long
With dlgColor
    ' No VB constant for CC_SOLIDCOLOR, but it works
    .Flags = cdlCCRGBInit Or CC_SOLIDCOLOR
    ' Make sure it's an RGB color
    .Color = TranslateColor(clr)
    .hWnd = hWnd
    ' Can only recognize cancel with error trapping
    .CancelError = True
    On Error Resume Next
    .ShowColor
    ' Return color, whether successful or not
    If Err Then
        OptionColor = clr
    Else
        OptionColor = .Color
    End If
End With
End Function
#End If

Sub OptionFont(Optional fSelection As Boolean)
    Dim f As Boolean, fnt As StdFont, clr As Long
    If fSelection Then
        Set fnt = New StdFont
        If IsNull(SelBold) Then fnt.Bold = fontDefault.Bold Else fnt.Bold = SelBold
        If IsNull(SelColor) Then clr = 0 Else clr = SelColor
        If IsNull(SelFontName) Then fnt.Name = fontDefault.Name Else fnt.Name = SelFontName
        If IsNull(SelFontSize) Then fnt.Size = fontDefault.Size Else fnt.Size = SelFontSize
        If IsNull(SelItalic) Then fnt.Italic = fontDefault.Italic Else fnt.Italic = SelItalic
        If IsNull(SelStrikeThru) Then fnt.Strikethrough = fontDefault.Strikethrough Else fnt.Strikethrough = SelStrikeThru
        If IsNull(SelUnderline) Then fnt.Underline = fontDefault.Underline Else fnt.Underline = SelUnderline
    Else
        Set fnt = Font
        clr = TextColor
    End If
    f = VBChooseFont(CurFont:=fnt, _
                     Color:=clr, _
                     Flags:=CF_EFFECTS Or CF_BOTH)
    If Not f Then Exit Sub
    If fSelection Then
        SelColor = clr
        SelBold = fnt.Bold
        SelFontName = fnt.Name
        SelFontSize = fnt.Size
        SelItalic = fnt.Italic
        SelStrikeThru = fnt.Strikethrough
        SelUnderline = fnt.Underline
    Else
        Set Font = fnt
        TextColor = clr
    End If
    Refresh
End Sub

Public Function DirtyDialog() As Boolean
    Dim s As String
    DirtyDialog = True ' Assume success
    ' Done if no dirty file to save
    If Not DirtyBit Then Exit Function
    ' Prompt for action if dirty file
    s = "File not saved: " & FileName & sCrLf & "Save now?"
    Select Case MsgBox(s, vbYesNoCancel)
    Case vbYes
        ' Save old file
        FileSave
    Case vbCancel
        ' User wants to terminate file change
        DirtyDialog = False
    Case vbNo
        ' Do nothing if user wants to throw away changes
    End Select
End Function

' Cut, copy, paste, and delete methods
Sub EditCopy()
    ' Copies text and formatting
    Call SendMessage(txt.hWnd, WM_COPY, ByVal 0&, ByVal 0&)
End Sub

Sub EditCut()
    ' Cuts/copies text and formatting
    Call SendMessage(txt.hWnd, WM_CUT, ByVal 0&, ByVal 0&)
End Sub

Sub EditPaste()
    ' Pastes text and formatting
    Call SendMessage(txt.hWnd, WM_PASTE, ByVal 0&, ByVal 0&)
End Sub

Sub EditDelete()
    ' Deletes text and formatting
    Call SendMessage(txt.hWnd, WM_CLEAR, ByVal 0&, ByVal 0&)
End Sub

Sub EditSelectAll()
    txt.SelStart = 0
    txt.SelLength = Me.Characters
End Sub

Sub EditUndo()
    Call SendMessage(txt.hWnd, EM_UNDO, ByVal 0&, ByVal 0&)
End Sub

Sub ClearUndo()
    Call SendMessage(txt.hWnd, EM_EMPTYUNDOBUFFER, ByVal 0&, ByVal 0&)
End Sub

Sub Scroll(Optional iLine As Long = 1, Optional iCol As Long = 0)
    SendMessage txt.hWnd, EM_LINESCROLL, ByVal iCol, ByVal iLine
End Sub

Sub ScrollToCaret()
    SendMessage txt.hWnd, EM_SCROLLCARET, ByVal 0&, ByVal 0&
End Sub

Sub PageUp()
    SendMessage txt.hWnd, WM_VSCROLL, ByVal SB_PAGEUP, ByVal 0&
End Sub

Sub PageDown()
    SendMessage txt.hWnd, WM_VSCROLL, ByVal SB_PAGEDOWN, ByVal 0&
End Sub

' Search and Replace functions

Sub SearchFind()
    ' Set properties on form
    Set finddlg.Editor = Me
    finddlg.ReplaceMode = False
    ' Load, but don't show yet
    Load finddlg
End Sub

Sub SearchFindNext()
    If FindWhat = sEmpty Then
        SearchFind
    Else
        Call FindNext(FindWhat)
    End If
End Sub

Sub SearchReplace()
    ' Set properties on form
    Set finddlg.Editor = Me
    finddlg.ReplaceMode = True
    ' Load, but don't show yet
    Load finddlg
End Sub

Function FindNext(Optional What As String, _
                  Optional MarkText As Boolean = True) As Integer
With txt
    Dim fWrap As Boolean, eso As ESearchOptions
    
    ' Set up options
    fWrap = (SearchOptionDirection = esdAll)
    If SearchOptionDirection = esdUp Then eso = esoBackward
    If SearchOptionCase Then eso = eso Or esoCaseSense
    If SearchOptionWord Then eso = eso Or esoWholeWord
    If What = sNullStr Then What = FindWhat
    
    ' Search for string
    Dim iPos As Integer
    If eso And esoBackward Then
        ' Adjust by one from end of search, then move back two
        iPos = .SelStart + .SelLength - 1  ' + 1 - 2
    Else
        ' Adjust by one then move forward one
        iPos = .SelStart + 2
    End If
    iPos = FindString(.Text, What, iPos, eso)

    ' If not found, wrap if appropriate
    If iPos = 0 Then
        If fWrap Then
            iPos = IIf(eso And esoBackward, Len(.Text), 1)
            iPos = FindString(.Text, What, iPos, eso)
        End If
    End If
    
    ' Mark found text if requested
    If MarkText And iPos Then
        .SelStart = iPos - 1
        .SelLength = Len(What)
    End If
    FindNext = iPos
End With
End Function

Function ReplaceNext(Optional Find As String, _
                     Optional Replace As String) As Integer
        
    If Find = sNullStr Then Find = FindWhat
    If Find = sEmpty Then Exit Function
    
    ' If first match not yet found, find it
    Dim i As Integer, s As String
    s = Mid(txt.Text, txt.SelStart + 1, Len(Find))
    If StrComp(s, Find, -SearchOptionCase) Then
        i = FindNext(Find)
        If i = 0 Then Exit Function
    Else ' If already found, make sure it's marked
        i = txt.SelStart + 1
        txt.SelLength = Len(Find)
    End If

    ' Replace text
    If i Then
        txt.SelText = Replace
        ReplaceNext = i
    End If

End Function

'' New properties
Property Get CanUndo() As Boolean
    CanUndo = SendMessage(txt.hWnd, EM_CANUNDO, ByVal 0&, ByVal 0&)
End Property

Property Get CanPaste() As Boolean
    If TextMode Then
        CanPaste = SendMessage(txt.hWnd, EM_CANPASTE, ByVal 0&, ByVal 0&)
    Else
        CanPaste = SendMessage(txt.hWnd, EM_CANPASTE, ByVal CF_TEXT, ByVal 0&)
    End If
End Property

Property Get TextMode() As Boolean
    TextMode = fTextMode
End Property

Property Let TextMode(ByVal fTextModeA As Boolean)
    ' Change to TextMode removes formatting
    If Not fTextMode And fTextMode <> fTextModeA Then
        ' Remove all formatting
        With txt
            Dim i As Long, c As Long
            i = .SelStart
            c = .SelLength
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelText = .Text
            .SelStart = i
            .SelLength = c
        End With
    End If
    fTextMode = fTextModeA
    PropertyChanged "TextMode"
End Property

' Optional argument character in which line is located
Property Get Line(Optional ByVal iChar As Long = -1) As Long
Attribute Line.VB_MemberFlags = "400"
    If iChar = -1 Then iChar = txt.SelStart
    ' Current line (zero adjusted)
    Line = 1 + GetLineFromChar(iChar)
End Property

Property Let Line(Optional ByVal iChar As Long = -1, ByVal iLineA As Long)
    ' Don't use optional parameter on Let
    BugAssert iChar = -1
    txt.SelStart = LinePosition(iLineA - 1) - 1
End Property

Property Get Lines() As Long
Attribute Lines.VB_MemberFlags = "400"
    ' Count of lines
    Lines = SendMessage(txt.hWnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
End Property

Property Get Character() As Long
Attribute Character.VB_MemberFlags = "400"
    ' Current character
    Character = txt.SelStart + 1
End Property

Property Let Character(ByVal iPos As Long)
    txt.SelStart = iPos - 1
End Property

Property Get Characters() As Long
Attribute Characters.VB_MemberFlags = "400"
    Dim i As Long
    ' Length is position of last line plus length of last line
    i = SendMessage(txt.hWnd, EM_LINEINDEX, ByVal Lines - 1, ByVal 0&)
    Characters = i + SendMessage(txt.hWnd, EM_LINELENGTH, _
                                 ByVal Lines - 1, ByVal 0&)
End Property

' Optional argument character in which column is located
Property Get Column(Optional ByVal iChar As Long = -1) As Long
Attribute Column.VB_MemberFlags = "400"
    If iChar = -1 Then iChar = txt.SelStart
    ' Column is current position minus position of line start
    Dim i As Long
    i = SendMessage(txt.hWnd, EM_LINEINDEX, _
                    ByVal Line(iChar) - 1, ByVal 0&)
    Column = Character - i
End Property

Property Let Column(Optional ByVal iChar As Long = -1, ByVal iColA As Long)
    ' Don't use optional parameter on Let
    BugAssert iChar = -1
    txt.SelStart = LinePosition + iColA - 2
End Property

Property Get Columns() As Long
Attribute Columns.VB_MemberFlags = "400"
    ' Column count is current line length
    Columns = SendMessage(txt.hWnd, EM_LINELENGTH, _
                          ByVal Character - 1, ByVal 0&)
End Property

Property Get Percent() As Integer
Attribute Percent.VB_MemberFlags = "400"
    Percent = (Character / (Characters + 1)) * 100
End Property

Property Let Percent(ByVal iA As Integer)
    txt.SelStart = Characters * (iA / 100)
End Property

Property Get LinePosition(Optional ByVal iLine As Long = -1) As Long
Attribute LinePosition.VB_MemberFlags = "400"
    LinePosition = SendMessage(txt.hWnd, EM_LINEINDEX, _
                               ByVal iLine, ByVal 0&) + 1
End Property

Property Get LineLength(Optional iLine As Long = -1) As Long
Attribute LineLength.VB_MemberFlags = "400"
    If iLine = -1 Then iLine = Line
    LineLength = SendMessage(txt.hWnd, EM_LINELENGTH, _
                             ByVal LinePosition(iLine), ByVal 0&)
End Property

Property Get FirstVisibleLine() As Long
    Line = 1 + SendMessage(txt.hWnd, EM_GETFIRSTVISIBLELINE, _
                           ByVal 0&, ByVal 0&)
End Property

Property Get LineText(Optional iLine As Long = -1) As String
Attribute LineText.VB_MemberFlags = "400"
    If iLine = -1 Then iLine = Line
    Const cCharMax = 252
    Dim s As String, c As Integer
    s = Space$(cCharMax + 3)
    Mid$(s, 1, 1) = Chr$(cCharMax And &HFF)
    Mid$(s, 2, 1) = Chr$(cCharMax \ 256)
    c = SendMessageStr(txt.hWnd, EM_GETLINE, iLine - 1, s)
    LineText = Left$(s, c)
End Property

Function GetLineFromChar(iPos As Long) As Long
Attribute GetLineFromChar.VB_Description = "Returns the number of the line containing a specified character position."
    GetLineFromChar = txt.GetLineFromChar(iPos)
End Function

Property Get Tabs() As Integer
Attribute Tabs.VB_MemberFlags = "400"
    Tabs = cCharPerTab
End Property

Property Let Tabs(ByVal cTab As Integer)
    Dim c As Long
    c = cTab * 4 ' Assume 4 dialog box units per character
    c = SendMessage(txt.hWnd, EM_SETTABSTOPS, ByVal 1&, c)
    cCharPerTab = cTab
    PropertyChanged "Tabs"
End Property

Property Get EnableTab() As Boolean
    EnableTab = fEnableTab
End Property

Property Let EnableTab(ByVal fEnableTabA As Boolean)
    fEnableTab = fEnableTabA
    PropertyChanged "EnableTab"
End Property


Sub SelVisible(fVisible As Boolean)
    SendMessage txt.hWnd, EM_HIDESELECTION, ByVal -fVisible, ByVal 0&
End Sub

' Find and Replace option properties
Property Get FindWhat(Optional iIndex As Long = 1) As String
Attribute FindWhat.VB_MemberFlags = "400"
With nFindWhat
    If .Count = 0 Or iIndex > .Count Then Exit Property
    FindWhat = .item(iIndex)
End With
End Property

Property Let FindWhat(Optional iIndex As Long = 1, sWhatA As String)
With nFindWhat
    ' Don't use optional parameter on Let
    BugAssert iIndex = 1
    Dim v As Variant, i As Long
    For i = 1 To .Count
        ' If item is in list, move to start of list
        If .item(i) = sWhatA Then
            .Add sWhatA, , 1
            .Remove i + 1
            NotifySearchChange eseFindWhat
            Exit Property
        End If
    Next
    ' If item isn't in list, add it
    If .Count Then
        .Add sWhatA, , 1
    Else
        .Add sWhatA
    End If
    NotifySearchChange eseFindWhat
End With
End Property

Property Get FindWhatCount() As Long
    FindWhatCount = nFindWhat.Count
End Property

Property Get FindWhatMax() As Long
    FindWhatMax = cFindWhatMax
End Property

Property Let FindWhatMax(cFindWhatMaxA As Long)
    cFindWhatMax = cFindWhatMaxA
    Dim v As Variant, i As Integer
    For i = nFindWhat.Count To cFindWhatMax + 1 Step -1
        ' If item is in list beyond maximum, remove it
        nFindWhat.Remove i
    Next
    NotifySearchChange eseFindWhat
End Property

Property Get ReplaceWith(Optional iIndex As Long = 1) As String
Attribute ReplaceWith.VB_MemberFlags = "400"
With nReplaceWith
    If .Count = 0 Or iIndex > .Count Then Exit Property
    ReplaceWith = .item(iIndex)
End With
End Property

Property Let ReplaceWith(Optional iIndex As Long = 1, sWithA As String)
With nReplaceWith
    ' Don't use optional parameter on Let
    BugAssert iIndex = 1
    Dim i As Integer ' i = 0
    For i = 1 To .Count
        ' If item is in list, move to start of list
        If .item(i) = sWithA Then
            .Add sWithA, , 1
            .Remove i + 1
            NotifySearchChange eseReplaceWith
            Exit Property
        End If
    Next
    ' If item isn't in list, add it
    If .Count Then
        .Add sWithA, , 1
    Else
        .Add sWithA
    End If
    NotifySearchChange eseReplaceWith
End With
End Property

Property Get ReplaceWithCount() As Long
    ReplaceWithMax = nReplaceWith.Count
End Property

Property Get ReplaceWithMax() As Long
    ReplaceWithMax = cReplaceWithMax
End Property

Property Let ReplaceWithMax(cReplaceWithMaxA As Long)
    cReplaceWithMax = cReplaceWithMaxA
    Dim i As Integer
    For i = cReplaceWithMax + 1 To nReplaceWith.Count
        ' If item is in list beyond maximum, remove it
        nReplaceWith.Remove i
    Next
    NotifySearchChange eseReplaceWith
End Property

Property Get SearchOptionDirection() As Integer
Attribute SearchOptionDirection.VB_MemberFlags = "400"
    SearchOptionDirection = esdDir
End Property

Property Let SearchOptionDirection(ByVal esdDirA As Integer)
    If esdDirA < 0 Or esdDirA > 2 Then esdDirA = 0
    esdDir = esdDirA
    NotifySearchChange eseDirection
End Property

Property Get SearchOptionCase() As Boolean
    SearchOptionCase = fSearchOptionCase
End Property

Property Let SearchOptionCase(ByVal fSearchOptionCaseA As Boolean)
    fSearchOptionCase = fSearchOptionCaseA
    NotifySearchChange eseCase
End Property

' Left as an exercise
Property Get SearchOptionWord() As Boolean
    SearchOptionWord = fSearchOptionWord
End Property

Property Let SearchOptionWord(ByVal fSearchOptionWordA As Boolean)
    fSearchOptionWord = fSearchOptionWordA
    NotifySearchChange eseWholeWord
End Property

Property Get SearchActive() As Boolean
    SearchActive = cSearchActive
End Property

' Friend so that only the search form can set this property
Friend Property Let SearchActive(ByVal fSearchActiveA As Boolean)
    ' Use reference count because you could have multiple search forms
    If fSearchActiveA Then
        cSearchActive = cSearchActive + 1
    Else
        cSearchActive = cSearchActive - 1
    End If
End Property

Property Get SaveWordWrap() As Boolean
    SaveWordWrap = fSaveWordWrap
End Property

Property Let SaveWordWrap(ByVal fSaveWordWrapA As Boolean)
    SendMessage txt.hWnd, EM_FMTLINES, ByVal -fSaveWordWrapA, ByVal 0&
    fSaveWordWrap = fSaveWordWrapA
    
    PropertyChanged "SaveWordWrap"
End Property

' RichTextBox properties passed through

Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets the paint style of a control at run time. "
Attribute Appearance.VB_MemberFlags = "400"
    Appearance = ordAppearance
End Property

Property Let Appearance(ByVal ordAppearanceA As AppearanceConstants)
    If Ambient.UserMode Then ErrRaise eeSetNotSupportedAtRuntime
    ModifyStyleBit txt.hWnd, ordAppearanceA = rtfThreeD, ES_SUNKEN
    ordAppearance = ordAppearanceA
    PropertyChanged "Appearance"
End Property

Property Get AutoVerbMenu() As Boolean
Attribute AutoVerbMenu.VB_Description = "Returns/sets a value that indicating whether the selected object's verbs will be displayed in a popup menu when the right mouse button is clicked."
    AutoVerbMenu = txt.AutoVerbMenu
End Property

Property Let AutoVerbMenu(ByVal fAutoVerbMenuA As Boolean)
    txt.AutoVerbMenu = fAutoVerbMenuA
    PropertyChanged "AutoVerbMenu"
End Property

Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of an object."
    BackColor = txt.BackColor
End Property

Property Let BackColor(ByVal clrBackColor As OLE_COLOR)
    txt.BackColor = clrBackColor
    PropertyChanged "BackColor"
End Property

Property Get BorderStyle() As RichTextLib.BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = txt.BorderStyle
End Property

Property Let BorderStyle(ByVal ordBorderStyle As RichTextLib.BorderStyleConstants)
    txt.BorderStyle = ordBorderStyle
    PropertyChanged "BorderStyle"
End Property

Property Get BulletIndent() As Single
Attribute BulletIndent.VB_Description = "Returns or sets the amount of indent used when SelBullet is set to True."
    BulletIndent = txt.BulletIndent
End Property

Property Let BulletIndent(ByVal rBulletIndentA As Single)
    If TextMode Then Exit Property
    txt.BulletIndent = rBulletIndentA
End Property

Property Get DirtyBit() As Boolean
Attribute DirtyBit.VB_MemberFlags = "400"
    DirtyBit = SendMessage(txt.hWnd, EM_GETMODIFY, ByVal 0&, ByVal 0&)
End Property

Property Let DirtyBit(ByVal fDirtyBitA As Boolean)
    Call SendMessage(txt.hWnd, EM_SETMODIFY, _
                     ByVal -CLng(fDirtyBitA), ByVal 0&)
    StatusEvent
End Property

#If 0 Then
Property Get DragIcon() As Picture
    Set DragIcon = txt.DragIcon
End Property

Property Let DragIcon(picDragIconA As Picture)
    Set txt.DragIcon = picDragIconA
    PropertyChanged "DragIcon"
End Property

Property Get DragMode() As DragModeConstants
    DragMode = txt.DragMode
End Property

Property Let DragMode(ByVal ordDragModeA As DragModeConstants)
    txt.DragMode = ordDragModeA
    PropertyChanged "DragMode"
End Property
#End If

' Read-only, run-time only
Property Get FilePath() As String
    '' Typo 'FileName' fixed in original
    FilePath = sFilePath
End Property

' Run time or design time
Property Get FileName() As String
    If sFilePath <> sEmpty Then FileName = GetFileBaseExt(sFilePath)
End Property

' Design-time only (use LoadFile at run time)
Property Let FileName(sFileNameA As String)
    If Ambient.UserMode Then ErrRaise eeSetNotSupportedAtRuntime
    ' Can't pass through design-time errors
    On Error GoTo FailFileName
    If sFileNameA = sEmpty Then
        ' Empty text only if it comes from a file
        If sFilePath <> sEmpty Then Text = sEmpty
        sFilePath = sEmpty
    Else
        sFileNameA = GetFullPath(sFileNameA)
        LoadFile sFileNameA
        sFilePath = sFileNameA
    End If
    PropertyChanged "FileName"
    Exit Property
FailFileName:
    ' Could empty FileName and Text, but I choose to ignore them
End Property

Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txt.Font
End Property

Property Set Font(ByVal fntA As Font)
    Dim fDirty As Boolean
    fDirty = DirtyBit
    Set txt.Font = fntA
    ' Changing the font shouldn't dirty the file in TextMode
    If TextMode Then DirtyBit = fDirty
    PropertyChanged "Font"
End Property

Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Returns/sets the foreground color used to display text."
Attribute TextColor.VB_MemberFlags = "400"
    TextColor = txt.SelColor
End Property

Property Let TextColor(ByVal clrTextColorA As OLE_COLOR)
With txt
    If TextMode Then
        Dim fEnabled As Boolean, fDirty As Boolean
        Dim iStart As Long, iLength As Long
        fDirty = DirtyBit
        ' Save selection
        SelVisible False
        iStart = .SelStart
        iLength = .SelLength
        ' Select all and change color
        .SelStart = 0
        .SelLength = Characters
        .SelColor = clrTextColorA
        ' Restore selection
        .SelStart = iStart
        .SelLength = iLength
        SelVisible True
        ' Changing the color shouldn't dirty the text in text mode
        DirtyBit = fDirty
    End If
    PropertyChanged "TextColor"
End With
End Property

Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value that specifies if the selected item remains highlighted when a control loses focus."
    HideSelection = txt.HideSelection
End Property

Property Let HideSelection(ByVal fHideSelectionA As Boolean)
    txt.HideSelection = fHideSelectionA
End Property

' Read only
Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = txt.hWnd
End Property

Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents can be edited."
    Locked = txt.Locked
End Property

Property Let Locked(ByVal fLockedA As Boolean)
    txt.Locked = fLockedA
    PropertyChanged "Locked"
End Property

Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets a value indicating the type of mouse pointer displayed when the mouse is over the control at run time."
    MousePointer = txt.MousePointer
End Property

Property Let MousePointer(ByVal ordMousePointerA As MousePointerConstants)
    txt.MousePointer = ordMousePointerA
    PropertyChanged "MousePointer"
End Property

Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = txt.MouseIcon
End Property

Property Set MouseIcon(ByVal picMouseIcon As Picture)
    Set txt.MouseIcon = picMouseIcon
    PropertyChanged "MouseIcon"
End Property

Property Get OLEDragMode() As OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this control can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = txt.OLEDragMode
End Property

Property Let OLEDragMode(ByVal ordOLEDragMode As OLEDragConstants)
    txt.OLEDragMode() = ordOLEDragMode
    PropertyChanged "OLEDragMode"
End Property

Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this control can act as an OLE drop target."
    OLEDropMode = txt.OLEDropMode
End Property

Property Let OLEDropMode(ByVal ordOLEDropMode As OLEDropConstants)
    txt.OLEDropMode() = ordOLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Property Get OLEObjects() As IOLEObjects
Attribute OLEObjects.VB_Description = "The insertable objects in an RTF file."
    Set OLEObjects = txt.OLEObjects
End Property

Property Get OverWrite() As Boolean
    OverWrite = fOverWrite
End Property

Property Let OverWrite(ByVal fOverWriteA As Boolean)
    ' Only change if value changed
    If fOverWriteA <> fOverWrite Then
        fOverWrite = fOverWriteA
        ' Change the keystate to match
        SendMessage txt.hWnd, WM_KEYDOWN, ByVal VK_INSERT, ByVal &H510001
        SendMessage txt.hWnd, WM_KEYUP, ByVal VK_INSERT, ByVal &HC0510001
        StatusEvent
        PropertyChanged "OverWrite"
    End If
End Property

Property Get ScrollBars() As ScrollBarsConstants
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether the control has horizontal or vertical scroll bars."
    ScrollBars = ordScrollBars
End Property

Property Let ScrollBars(ordScrollBarsA As ScrollBarsConstants)
    Dim af As Long, hParent As Long
    Const afMask As Long = Not (WS_HSCROLL Or WS_VSCROLL)
    af = GetWindowLong(txt.hWnd, GWL_STYLE) And afMask
    
    ' Call InitScrollBars once for the lifetime of control
    Static fNotFirstTime As Boolean
    If fNotFirstTime = False Then
        InitScrollBars af
        fNotFirstTime = True
    End If
    
    Select Case ordScrollBarsA
    Case rtfNone
        ' Done
    Case rtfHorizontal
        af = af Or WS_HSCROLL
    Case rtfVertical
        af = af Or WS_VSCROLL
    Case rtfBoth
        af = af Or WS_HSCROLL Or WS_VSCROLL
    End Select
    Call SetWindowLong(hWnd, GWL_STYLE, af)
    ' Reset the parent so that change will "take"
    hParent = GetParent(hWnd)
    SetParent hWnd, hParent
    ' Redraw for added insurance
    Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
                      SWP_NOZORDER Or SWP_NOSIZE Or _
                      SWP_NOMOVE Or SWP_DRAWFRAME)

    ordScrollBars = ordScrollBarsA
    PropertyChanged "ScrollBars"
End Property

Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Sets the right margin used for textwrap, centering, etc."
    RightMargin = ScaleX(txt.RightMargin, vbTwips, Parent.ScaleMode)
End Property

Property Let RightMargin(ByVal rRightMargin As Single)
    If rRightMargin < 0 Or rRightMargin > 65535 Then
        rRightMargin = 65535
    End If
    txt.RightMargin = ScaleX(rRightMargin, Parent.ScaleMode, vbTwips)
    PropertyChanged "RightMargin"
End Property

Property Get SelAlignment() As Variant
Attribute SelAlignment.VB_Description = "Returns/sets a value that controls the alignment of the paragraphs."
Attribute SelAlignment.VB_MemberFlags = "400"
    If TextMode Then
        SelAlignment = Null
    Else
        SelAlignment = txt.SelAlignment
    End If
End Property

Property Let SelAlignment(ByVal vSelAlignmentA As Variant)
    If TextMode Then Exit Property
    txt.SelAlignment = vSelAlignmentA
End Property

Property Get SelBold() As Variant
Attribute SelBold.VB_Description = "Returns/set the bold format of the currently selected text."
Attribute SelBold.VB_MemberFlags = "400"
    If TextMode Then
        SelBold = Null
    Else
        SelBold = txt.SelBold
    End If
End Property

Property Let SelBold(ByVal vSelBold As Variant)
    If TextMode Then Exit Property
    txt.SelBold = vSelBold
End Property

Property Get SelBullet() As Variant
Attribute SelBullet.VB_Description = "Returns/sets a value that determines if a paragraph in the control containing the current selection or insertion point has the bullet style."
Attribute SelBullet.VB_MemberFlags = "400"
    If TextMode Then
        SelBullet = Null
    Else
        SelBullet = txt.SelBullet
    End If
End Property

Property Let SelBullet(ByVal vSelBullet As Variant)
    If TextMode Then Exit Property
    txt.SelBullet = vSelBullet
End Property

Property Get SelCharOffset() As Variant
Attribute SelCharOffset.VB_Description = "Returns/sets a value that determines whether text in the control appears on the baseline (normal), as a superscript above the baseline, or as a subscript below the baseline."
Attribute SelCharOffset.VB_MemberFlags = "400"
    If TextMode Then
        SelCharOffset = Null
    Else
        SelCharOffset = txt.SelCharOffset
    End If
End Property

Property Let SelCharOffset(ByVal vSelCharOffset As Variant)
    If TextMode Then Exit Property
    txt.SelCharOffset = vSelCharOffset
End Property

Property Get SelColor() As Variant
Attribute SelColor.VB_Description = "Returns/sets a value that determines the color of text in the control."
Attribute SelColor.VB_MemberFlags = "400"
    If TextMode Then
        SelColor = Null
    Else
        SelColor = txt.SelColor
    End If
End Property

Property Let SelColor(ByVal vSelColor As Variant)
    If TextMode Then Exit Property
    txt.SelColor = vSelColor
End Property

Property Get SelFontName() As Variant
Attribute SelFontName.VB_Description = "Returns/sets the font used to display the currently selected text or the character(s) immediately following the insertion point in the control."
Attribute SelFontName.VB_MemberFlags = "400"
    If TextMode Then
        SelFontName = Null
    Else
        SelFontName = txt.SelFontName
    End If
End Property

Property Let SelFontName(ByVal vSelFontName As Variant)
    If TextMode Then Exit Property
    txt.SelFontName = vSelFontName
End Property

Property Get SelFontSize() As Variant
Attribute SelFontSize.VB_Description = "Returns/sets a value that specifies the size of the font used to display text."
Attribute SelFontSize.VB_MemberFlags = "400"
    If TextMode Then
        SelFontSize = Null
    Else
        SelFontSize = txt.SelFontSize
    End If
End Property

Property Let SelFontSize(ByVal vSelFontSize As Variant)
    If TextMode Then Exit Property
    txt.SelFontSize = vSelFontSize
End Property

Property Get SelHangingIndent() As Variant
Attribute SelHangingIndent.VB_Description = "Returns/sets the distance between left edge of the first line of text in the selected paragraph(s) (as specified by SelIndent) and the left edge of subsequent lines of text in the same paragraphs."
Attribute SelHangingIndent.VB_MemberFlags = "400"
    If TextMode Then
        SelHangingIndent = Null
    Else
        SelHangingIndent = txt.SelHangingIndent
    End If
End Property

Property Let SelHangingIndent(ByVal vSelHangingIndent As Variant)
    If TextMode Then Exit Property
    txt.SelHangingIndent = vSelHangingIndent
End Property

Property Get SelIndent() As Variant
Attribute SelIndent.VB_Description = "Returns/sets the distance between the left edge of the control and the left edge of the text that is selected or added at the current insertion point."
Attribute SelIndent.VB_MemberFlags = "400"
    If TextMode Then
        SelIndent = Null
    Else
        SelIndent = txt.SelIndent
    End If
End Property

Property Let SelIndent(ByVal vSelIndent As Variant)
    If TextMode Then Exit Property
    txt.SelIndent = vSelIndent
End Property

Property Get SelItalic() As Variant
Attribute SelItalic.VB_Description = "Returns/set the italic format of the currently selected text."
Attribute SelItalic.VB_MemberFlags = "400"
    If TextMode Then
        SelItalic = Null
    Else
        SelItalic = txt.SelItalic
    End If
End Property

Property Let SelItalic(ByVal vSelItalic As Variant)
    If TextMode Then Exit Property
    txt.SelItalic = vSelItalic
End Property

Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = txt.SelLength
End Property

Property Let SelLength(ByVal iSelLengthA As Long)
    txt.SelLength = iSelLengthA
End Property

Property Get SelProtected() As Variant
Attribute SelProtected.VB_Description = "Returns/sets a value that determines if the selected text is protected against editing."
Attribute SelProtected.VB_MemberFlags = "400"
    If TextMode Then
        SelProtected = Null
    Else
        SelProtected = txt.SelProtected
    End If
End Property

Property Let SelProtected(ByVal vSelProtected As Variant)
    If TextMode Then Exit Property
    txt.SelProtected = vSelProtected
End Property

Property Get SelRightIndent() As Variant
Attribute SelRightIndent.VB_Description = "Returns/sets the distance between the right edge of the control and the right edge of the text that is selected or added at the current insertion point."
Attribute SelRightIndent.VB_MemberFlags = "400"
    If TextMode Then
        SelRightIndent = Null
    Else
        SelRightIndent = txt.SelRightIndent
    End If
End Property

Property Let SelRightIndent(ByVal vSelRightIndent As Variant)
    If TextMode Then Exit Property
    txt.SelRightIndent = vSelRightIndent
End Property

Property Get SelRTF() As String
Attribute SelRTF.VB_Description = "Returns/sets the text (in .RTF format) in the current selection of a control."
Attribute SelRTF.VB_MemberFlags = "400"
    If TextMode Then Exit Property
    SelRTF = txt.SelRTF
End Property

Property Let SelRTF(sSelRTFA As String)
    If TextMode Then Exit Property
    txt.SelRTF = sSelRTFA
End Property

Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = txt.SelStart
End Property

Property Let SelStart(ByVal iSelStartA As Long)
    txt.SelStart = iSelStartA
End Property

Property Get SelStrikeThru() As Variant
Attribute SelStrikeThru.VB_Description = "Returns/set the strikethru format of the currently selected text."
Attribute SelStrikeThru.VB_MemberFlags = "400"
    If TextMode Then
        SelStrikeThru = Null
    Else
        SelStrikeThru = txt.SelStrikeThru
    End If
End Property

Property Let SelStrikeThru(ByVal vSelStrikeThru As Variant)
    If TextMode Then Exit Property
    txt.SelStrikeThru = vSelStrikeThru
End Property

Property Get SelTabCount() As Variant
Attribute SelTabCount.VB_Description = "Returns/sets the number of tabs.  Used in conjunction with the SelTab Property."
Attribute SelTabCount.VB_MemberFlags = "400"
    If TextMode Then
        SelTabCount = Null
    Else
        SelTabCount = txt.SelTabCount
    End If
End Property

Property Let SelTabCount(ByVal vSelTabCount As Variant)
    If TextMode Then Exit Property
    txt.SelTabCount = vSelTabCount
End Property

Property Get SelTabs(iElement As Integer) As Variant
Attribute SelTabs.VB_Description = "Returns/sets the absolute tab positions of text.  Used in conjunction with the SelTabCount Property."
Attribute SelTabs.VB_MemberFlags = "400"
    If TextMode Then
        SelTabs = Null
    Else
        SelTabs = ScaleX(txt.SelTabs(iElement), vbTwips, Parent.ScaleMode)
        'SelTabs = txt.SelTabs(iElement)
    End If
End Property

Property Let SelTabs(iElement As Integer, ByVal vSelTabs As Variant)
    If TextMode Then Exit Property
    txt.SelTabs(iElement) = ScaleX(vSelTabs, Parent.ScaleMode, vbTwips)
    'txt.SelTabs(iElement) = vSelTabs
End Property

Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text; consists of a zero-length string if no characters are selected."
Attribute SelText.VB_MemberFlags = "400"
    SelText = txt.SelText
End Property

Property Let SelText(sSelTextA As String)
    txt.SelText = sSelTextA
End Property

Property Get SelUnderline() As Variant
Attribute SelUnderline.VB_Description = "Returns/set the underline format of the currently selected text."
Attribute SelUnderline.VB_MemberFlags = "400"
    If TextMode Then
        SelUnderline = Null
    Else
        SelUnderline = txt.SelUnderline
    End If
End Property

Property Let SelUnderline(ByVal vSelUnderline As Variant)
    If TextMode Then Exit Property
    txt.SelUnderline = vSelUnderline
End Property

Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_UserMemId = 0
    If Ambient.UserMode Then
        Text = txt.Text
    Else
        ' Show only the first line in property page
        Dim iPos As Long
        iPos = InStr(txt.Text, sCr)
        If iPos Then
            Text = Left$(txt.Text, iPos - 1) & "..."
        Else
            Text = txt.Text
        End If
    End If
End Property

Property Let Text(sTextA As String)
    txt.Text = sTextA
    PropertyChanged "Text"
End Property

Property Get TextRTF() As String
Attribute TextRTF.VB_Description = "Returns/sets the text of the control, including all .RTF code."
Attribute TextRTF.VB_MemberFlags = "400"
    If TextMode Then
        TextRTF = txt.TextRTF
    Else
        TextRTF = sEmpty
    End If
End Property

Property Let TextRTF(sTextRTFA As String)
    If fTextMode = rtfRTF Then
        TextRTF = txt.TextRTF
    ' Else ignore
    End If
End Property

Property Get ScaleLeft() As Single
    Dim rc As RECT
    SendMessage txt.hWnd, EM_GETRECT, ByVal 0&, rc
    ScaleLeft = Extender.Left + (rc.Left * Screen.TwipsPerPixelY)
End Property

Property Get ScaleTop() As Single
    Dim rc As RECT
    SendMessage txt.hWnd, EM_GETRECT, ByVal 0&, rc
    ScaleTop = Extender.Top + (rc.Top * Screen.TwipsPerPixelY)
End Property

Property Get ScaleWidth() As Single
    Dim rc As RECT
    SendMessage txt.hWnd, EM_GETRECT, ByVal 0&, rc
    ScaleWidth = (rc.Right - rc.Left) * Screen.TwipsPerPixelX
End Property

Property Get ScaleHeight() As Single
    Dim rc As RECT
    SendMessage txt.hWnd, EM_GETRECT, ByVal 0&, rc
    ScaleHeight = (rc.bottom - rc.Top) * Screen.TwipsPerPixelY
End Property

Property Get LeftMargin() As Single
    ' Commented out because property has no effect for reasons unknown
    'Dim dx As Long
    'dx = SendMessage(txt.hWnd, EM_GETMARGINS, ByVal 0&, ByVal 0&)
    'LeftMargin = ScaleX(LoWord(dx), vbTwips, vbPixels)
End Property

Property Let LeftMargin(ByVal rLeftMargin As Single)
    ' Commented out because property has no effect for reasons unknown
    'Dim dx As Long
    'Const EC_LEFTMARGIN = 1
    'dx = ScaleX(CLng(rLeftMargin), vbPixels, vbTwips)
    'SendMessage txt.hWnd, EM_SETMARGINS, ByVal EC_LEFTMARGIN, ByVal dx
End Property

'' Methods passed through

Public Sub Drag(Optional ByVal ordAction As Integer = vbBeginDrag)
    txt.Drag ordAction
End Sub

' Pass through txt method, but it's similar to FindNext
Function Find(sSearch As String, Optional vStart As Variant, _
              Optional vEnd As Variant, _
              Optional afOptions As Integer = 0) As Long
Attribute Find.VB_Description = "Searches the text for a given string."
With txt
    If IsMissing(vStart) Then
        If IsMissing(vEnd) Then
            ' Both missing
            If txt.SelLength > 0 Then
                vStart = .SelStart
                vEnd = .SelStart + .SelLength
            Else
                ' Enhance to start at current position
                vStart = .SelStart
                vEnd = .SelStart - 1
            End If
        Else
            ' Start missing
            vStart = .SelStart
        End If
    Else
        If IsMissing(vEnd) Then
            ' End missing
            vEnd = Characters
        ' else
            ' None missing
        End If
    End If
    .Find sSearch, vStart, vEnd, afOptions
End With
End Function

' Run-time only (use FileName at design time)
Sub LoadFile(sFileNameA As String, _
             Optional ordTextModeA As ELoadSave = elsDefault)
Attribute LoadFile.VB_Description = "Loads an .RTF file or text file."
    If sFileNameA = sEmpty Then Exit Sub
    BugAssert ordTextModeA >= elsDefault And ordTextModeA <= elstext
    If ordTextModeA = elsDefault Then
        ordTextModeA = IIf(TextMode, elstext, elsrtf)
    End If
    If TextMode Then Set Font = fontDefault
    ' Don't reload clean file
    sFileNameA = GetFullPath(sFileNameA)
    If sFileNameA = sFilePath And DirtyBit = False Then Exit Sub
    ' Use RichTextBox method (raise unhandled errors to caller)
    txt.LoadFile sFileNameA, ordTextModeA
    sFilePath = sFileNameA
    DirtyBit = False
End Sub

Sub Move(x As Single, Optional y As Variant, _
         Optional dx As Variant, Optional dy As Variant)
    If IsMissing(y) Then
        txt.Move x
    ElseIf IsMissing(dx) Then
        txt.Move x, y
    ElseIf IsMissing(dy) Then
        txt.Move x, y, dx
    Else
        txt.Move x, y, dx, dy
    End If
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    txt.OLEDrag
End Sub

Public Sub SelPrint(ByVal hDC As Long, Optional fPrintAll As Boolean)
Attribute SelPrint.VB_Description = "Sends formatted text to a device for printing."
With txt
    If fPrintAll Then
        Dim c As Long
        c = .SelLength
        .SelLength = 0
        .SelPrint hDC
        .SelLength = c
    Else
        .SelPrint hDC
    End If
End With
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint."
    txt.Refresh
End Sub

Public Sub SetFocus()
    txt.SetFocus
End Sub

Sub SaveFile(sFileNameA As String, _
             Optional ordTextModeA As ELoadSave = elsDefault)
    If sFileNameA = sEmpty Then Exit Sub
    BugAssert ordTextModeA >= elsDefault And ordTextModeA <= elstext
    If ordTextModeA = elsDefault Then
        ordTextModeA = IIf(TextMode, elstext, elsrtf)
    End If
    ' Use RichTextBox method (raise unhandled errors to caller)
    sFileNameA = GetFullPath(sFileNameA)
    txt.SaveFile sFileNameA, ordTextModeA
    sFilePath = sFileNameA
    DirtyBit = False
End Sub

Sub ShowWhatsThis()
    txt.ShowWhatsThis
End Sub

Public Sub Span(sCharSet As String, _
                Optional fForward As Boolean = True, _
                Optional fNegate As Boolean = False)
    txt.Span sCharSet, fForward, fNegate
End Sub

Public Sub UpTo(sCharSet As String, _
                Optional fForward As Boolean = True, _
                Optional fNegate As Boolean = False)
Attribute UpTo.VB_Description = "Moves the insertion point up to, but not including, the first character that is a member of the specified character set."
    txt.UpTo sCharSet, fForward, fNegate
End Sub

' Events
Sub txt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Sub txt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Sub txt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyInsert Then
        ' Insert state changed, so change the variable to match
        fOverWrite = Not fOverWrite
        StatusEvent
    End If
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Sub txt_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Sub txt_DblClick()
    RaiseEvent DblClick
End Sub

Sub txt_Click()
    RaiseEvent Click
End Sub

Sub txt_Change()
    RaiseEvent Change
End Sub

Private Sub txt_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub txt_OLEDragDrop(data As RichTextLib.DataObject, Effect As Long, _
                            Button As Integer, Shift As Integer, _
                            x As Single, y As Single)
    RaiseEvent OLEDragDrop(data, Effect, Button, Shift, x, y)
End Sub

Private Sub txt_OLEDragOver(data As RichTextLib.DataObject, Effect As Long, _
                            Button As Integer, Shift As Integer, _
                            x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub txt_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub txt_OLESetData(data As RichTextLib.DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(data, DataFormat)
End Sub

Private Sub txt_OLEStartDrag(data As RichTextLib.DataObject, _
                             AllowedEffects As Long)
    RaiseEvent OLEStartDrag(data, AllowedEffects)
End Sub

Private Sub txt_SelChange()
    RaiseEvent SelChange
    StatusEvent
End Sub

Private Sub StatusEvent()
    Dim iLine As Long, cLine As Long, iCol As Long, i As Long
    Dim cCol As Long, iChar As Long, cChar As Long
        
    ' Count of lines
    cLine = SendMessage(txt.hWnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
    ' Current line (zero adjusted)
    iLine = 1 + txt.GetLineFromChar(txt.SelStart)
    ' Current character
    iChar = txt.SelStart + 1
    ' Length is position of last line plus length of last line
    cChar = SendMessage(txt.hWnd, EM_LINEINDEX, ByVal cLine - 1, ByVal 0&)
    ' Bug fix 1/9/98
    ' User report of this condition; I couldn't duplicate, but prevented
    If cChar = -1 Then cChar = 0
    i = SendMessage(txt.hWnd, EM_LINELENGTH, ByVal cChar - 1, ByVal 0&)
    cChar = cChar + i
    ' Column count is current line length
    cCol = SendMessage(txt.hWnd, EM_LINELENGTH, ByVal iChar - 1, ByVal 0&)
    ' Column is current position minus position of line start
    i = SendMessage(txt.hWnd, EM_LINEINDEX, ByVal iLine - 1, ByVal 0&)
    iCol = iChar - i
    RaiseEvent StatusChange(iLine, cLine, iCol, cCol, iChar, cChar, DirtyBit)
End Sub

'' Private helpers
Private Function FilterString() As String
    Dim s As String, v As Variant
    For Each v In nFilters
        s = s & v & "|"
    Next
    FilterString = Left$(s, Len(s) - 1)
End Function

Private Sub StringFilter(sFilters As String)
    '
End Sub

Private Sub InitFilters()
With nFilters
    If .Count = 0 Then
        .Add "Text files (*.txt): *.txt"
        .Add "Rich text files (*.rtf): *.rtf"
        .Add "All files (*.*): *.*"
    End If
End With
End Sub

' For reasons unknown, the rich text box ScrollBars property
' misbehaves at design time. InitScrollBars hacks around this
' problem by setting horizontal and vertical scroll bars in
' succession, after which the ScrollBars property works correctly.
Private Sub InitScrollBars(af As Long)
    Dim hParent As Long
    
    ' Set horizontal scroll bar
    Call SetWindowLong(hWnd, GWL_STYLE, af Or WS_HSCROLL)
    ' Reset the parent so that change will "take"
    hParent = GetParent(hWnd)
    SetParent hWnd, hParent
    ' Redraw for added insurance
    Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
                      SWP_NOZORDER Or SWP_NOSIZE Or _
                      SWP_NOMOVE Or SWP_DRAWFRAME)
                      
    ' Set vertical scroll bar
    Call SetWindowLong(hWnd, GWL_STYLE, af Or WS_VSCROLL)
    ' Reset the parent so that change will "take"
    hParent = GetParent(hWnd)
    SetParent hWnd, hParent
    ' Redraw for added insurance
    Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
                      SWP_NOZORDER Or SWP_NOSIZE Or _
                      SWP_NOMOVE Or SWP_DRAWFRAME)
End Sub

Sub NotifySearchChange(ByVal ese As ESearchEvent)
    RaiseEvent SearchChange(ese)
    If SearchActive Then finddlg.SearchChange ese
End Sub



