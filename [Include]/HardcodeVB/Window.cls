VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private hWnd As Long

Function CreateFromHandle(hWndA) As Boolean
    If IsWindow(hWndA) Then hWnd = hWndA Else hWnd = hNull
End Function

Function CreateFromPoint(x As Long, y As Long) As Boolean
    hWnd = WindowFromPoint(x, y)
    CreateFromPoint = (hWnd <> hNull)
End Function

Function CreateFromFind(Optional sClass As String = vbNullString, _
                        Optional sTitle As String = vbNullString) As Boolean
    hWnd = FindWindow(sClass, sTitle)
    CreateFromFind = (hWnd = hNull)
End Function

Function CreateFromActive() As Boolean
    hWnd = GetActiveWindow()
    CreateFromActive = (hWnd = hNull)
End Function

Function CreateFromFocus() As Boolean
    hWnd = GetFocus()
    CreateFromFocus = (hWnd = hNull)
End Function

' SendMessage

' IsChild
' IsVisible
' IsEnabled

' Show
' Enable

' ClientRect
' WindowRect
' TopWindow

' Get/SetWindowPlacement
' SetWindowPos
' Move
' BringToTop
' Flash
' ClientToScreen
' ScreenToClient
' MapPoints
' Update
' MakeActive
' MakeFocus
' MakeSysModal
' Destroy
' GetLastPopup
' StartCapture
' EndCapture

' Menu
' SystemMenu
' GiveMenu

Sub Capture(f As Boolean)
    If f Then
        SetCapture hWnd
    Else
        ReleaseCapture
    End If
End Sub

''@B Handle
Public Property Get Handle() As Long
Attribute Handle.VB_UserMemId = 0
    Handle = hWnd
End Property

Public Property Let Handle(ByVal hWndA As Long)
    If IsWindow(hWndA) Then hWnd = hWndA Else hWnd = hNull
End Property
''@E Handle

Property Get Instance() As Long
    Instance = GetWindowLong(hWnd, GWL_HINSTANCE)
End Property

Property Get Module() As Long
    Module = GetClassLong(hWnd, GCL_HMODULE)
End Property

Property Get Process() As Long
    Dim id As Long
    GetWindowThreadProcessId hWnd, id
    Process = OpenProcess(PROCESS_ALL_ACCESS, False, id)
End Property

Property Get Style() As Long
    Style = GetWindowLong(hWnd, GWL_STYLE)
End Property

Property Get ExStyle() As Long
    ExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
End Property

Property Get ClassStyle() As Long
    ClassStyle = GetClassLong(hWnd, GCL_STYLE)
End Property

Property Get Parent() As Long
    Parent = GetParent(hWnd)
End Property

Function NewParent(hWndNew As Long) As Long
    NewParent = SetParent(hWnd, hWndNew)
End Function

Property Get Owner() As Long
    Owner = GetWindow(hWnd, GW_OWNER)
End Property

Property Get Child() As Long
    Child = GetWindow(hWnd, GW_CHILD)
End Property

Property Get First() As Long
    First = GetWindow(hWnd, GW_HWNDFIRST)
End Property

Property Get Last() As Long
    Last = GetWindow(hWnd, GW_HWNDLAST)
End Property

Property Get NextWin() As Long
    NextWin = GetWindow(hWnd, GW_HWNDNEXT)
End Property

Property Get Previous() As Long
    Previous = GetWindow(hWnd, GW_HWNDPREV)
End Property

''@B ExeName
Property Get ExeName() As String
    Dim sName As String, cName As Long, i As Long

    sName = String$(cMaxPath + 1, 0)
    cName = GetModuleFileName(Instance, sName, cMaxPath + 1)
#If afDebug Then
    If cName = 0 Then
         BugMessage "Error: " & Err.LastDllError
    End If
#End If
    sName = Left$(sName, cName)
    For i = Len(sName) To 1 Step -1
        If Mid$(sName, i, 1) = "\" Then Exit For
    Next
    ExeName = Mid$(sName, i + 1)
End Property
''@E ExeName

''@B ExePath
Property Get ExePath() As String
    Dim sName As String, cName As Integer

    sName = String$(cMaxPath + 1, 0)
    cName = GetModuleFileName(Instance, sName, cMaxPath)
    ExePath = Left$(sName, cName)
#If afDebug Then
    If cName = 0 Then
        BugMessage "Error: " & Err.LastDllError
    End If
#End If
End Property
''@E ExePath

''@B Class
Property Get ClassName() As String
    Dim sName As String, cName As Integer
    sName = String$(41, 0)
    cName = GetClassName(hWnd, sName, 41)
    ClassName = Left$(sName, cName)
End Property
''@E Class

''@B Owner
Property Get OwnerName() As String
    Dim hwndOwner As Long
    hwndOwner = GetWindow(hWnd, GW_OWNER)
    If hwndOwner <> hNull Then
        Dim wndOwner As New CWindow
        wndOwner.CreateFromHandle hwndOwner
        OwnerName = wndOwner.Caption
    Else
        OwnerName = sEmpty
    End If
End Property
''@E Owner

''@B Caption
Property Get Caption() As String
    Dim cCaption As Integer, sCaption As String
    sCaption = Space$(81)
    cCaption = GetWindowText(hWnd, sCaption, 81)
    Caption = Left$(sCaption, cCaption)
End Property
''@E Caption

Property Let Caption(sCaptionA As String)
    SetWindowText hWnd, sCaptionA
End Property

''@B CaptionLine
Property Get CaptionLine() As String
    Dim cCaption As Integer, sCaption As String
    sCaption = Space$(81)
    cCaption = GetWindowText(hWnd, sCaption, 81)
    sCaption = Left$(sCaption, cCaption)
    ' Chop off end of multi-line captions
    cCaption = InStr(sCaption, sCr)
    CaptionLine = IIf(cCaption, Left$(sCaption, cCaption), sCaption)
End Property
''@E CaptionLine
