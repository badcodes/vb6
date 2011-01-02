VERSION 5.00
Begin VB.UserControl ctlSplitterEx 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   ScaleHeight     =   3600
   ScaleWidth      =   6105
   Begin VB.Label lblInfo 
      Caption         =   "(must be behind its attached objects)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label lblInfo 
      Caption         =   "SplitterEx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "ctlSplitterEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum tTILE_MODE
    TILE_HORIZONTALLY
    TILE_VERTICALLY
End Enum

Dim m_oMe As Object
Dim m_oLeft As Object
Dim m_oRight As Object
Dim m_bInitializes As Boolean
Dim m_lTilePercence As Long
Dim m_tTileMode As tTILE_MODE

Dim m_bMoving As Boolean

Dim m_MinTilePercent As Long
Dim m_MaxTilePercent As Long
Const TILE_BAR_SIZE = 25

Public Event Resized()


Public Property Get HWND() As Long
HWND = UserControl.HWND
End Property
Public Property Let MaxTilePercent(mt As Long)
If mt > 5 And mt < 98 And mt > m_MinTilePercent Then m_MaxTilePercent = mt
End Property
Public Property Let MinTilePercent(mt As Long)
If mt > 5 And mt < 98 And mt < m_MaxTilePercent Then m_MinTilePercent = mt
End Property

Public Property Let TileMode(tm As tTILE_MODE)
m_tTileMode = tm
Refresh
End Property

Private Sub FindOwnInstance()
Dim i As Long

Dim s As String, s2 As String
s = UserControl.Name
Dim Myhwnd As Long
Myhwnd = UserControl.HWND

For i = 0 To UserControl.ParentControls.count - 1
    If TypeOf UserControl.ParentControls.Item(i) Is ctlSplitterEx Then
            If UserControl.ParentControls.Item(i).HWND = Myhwnd Then
                Set m_oMe = UserControl.ParentControls.Item(i)
                Exit For
            End If
    End If
Next

Exit Sub

For i = 0 To UserControl.ParentControls.count - 1
    If TypeOf UserControl.ParentControls.Item(i) Is ctlSplitterEx Then
        Set m_oMe = UserControl.ParentControls.Item(i)
        Exit For
    End If
Next

End Sub

Private Sub InternalInit()
m_MinTilePercent = 5
m_MaxTilePercent = 98
m_lTilePercence = 30
lblInfo(1).Visible = False
lblInfo(0).Visible = False
FindOwnInstance
m_bInitializes = True
End Sub
Public Sub AttachObjects(oLeft As Object, oRight As Object, Optional InitZOrder As Boolean = False)
Set m_oLeft = oLeft
Set m_oRight = oRight
InternalInit
If InitZOrder = True Then
    oLeft.ZOrder    'to front
    oRight.ZOrder   'to front
End If
Refresh
End Sub
Public Sub Refresh()
ResizeObjects
End Sub
Private Sub ResizeObjects()
On Error Resume Next
If m_tTileMode = TILE_VERTICALLY Then
    '------ Linkes Objekt
    m_oLeft.Top = m_oMe.Top
    m_oLeft.Height = m_oMe.Height
    m_oLeft.Left = m_oMe.Left
    
    Dim lLeftWidth As Long
    lLeftWidth = (m_oMe.Width * m_lTilePercence) / 100
    m_oLeft.Width = lLeftWidth
    
    '------ Rechtes Objekt
    m_oRight.Top = m_oMe.Top
    m_oRight.Height = m_oMe.Height
    
    m_oRight.Left = m_oLeft.Left + m_oLeft.Width + TILE_BAR_SIZE
    m_oRight.Width = (m_oMe.Left + m_oMe.Width) - (m_oLeft.Left + m_oLeft.Width + TILE_BAR_SIZE)
Else
    '------ Oberes Objekt
    m_oLeft.Top = m_oMe.Top
    m_oLeft.Width = m_oMe.Width
    m_oLeft.Left = m_oMe.Left

    Dim lLeftHeight As Long
    lLeftHeight = (m_oMe.Height * m_lTilePercence) / 100
    m_oLeft.Height = lLeftHeight
           
    '------ Unteres Objekt
    m_oRight.Left = m_oMe.Left
    m_oRight.Width = m_oMe.Width
    m_oRight.Top = m_oLeft.Top + m_oLeft.Height + TILE_BAR_SIZE
    m_oRight.Height = (m_oMe.Top + m_oMe.Height) - (m_oLeft.Top + m_oLeft.Height + TILE_BAR_SIZE)
End If

If TypeOf m_oLeft Is ctlSplitterEx Then m_oLeft.Refresh
If TypeOf m_oRight Is ctlSplitterEx Then m_oLeft.Refresh

RaiseEvent Resized
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TileStart As Long
Dim MousePos As Long

If m_tTileMode = TILE_VERTICALLY Then
    TileStart = m_oLeft.Width
    MousePos = X
Else
    TileStart = m_oLeft.Height
    MousePos = Y
End If
  
If MousePos > TileStart And MousePos < TileStart + TILE_BAR_SIZE Or m_bMoving = True Then
    m_bMoving = True
Else
    m_bMoving = False
End If



End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TileStart As Long
Dim MousePos As Long
Dim Relative As Long

If m_tTileMode = TILE_VERTICALLY Then
    TileStart = m_oLeft.Width
    Relative = m_oMe.Width
    MousePos = X
Else
    TileStart = m_oLeft.Height
    Relative = m_oMe.Height
    MousePos = Y
End If


If MousePos > TileStart And MousePos < TileStart + TILE_BAR_SIZE Or m_bMoving = True Then
    UserControl.MousePointer = IIf(m_tTileMode = TILE_VERTICALLY, vbSizeWE, vbSizeNS)
Else
    UserControl.MousePointer = vbDefault
End If


If m_bMoving = True Then
    ' Calculate Current Delta
    Dim pc As Long
    'pc = 100 / m_oMe.Width * X
    pc = 100 / Relative * MousePos
    If pc > m_MinTilePercent And pc < m_MaxTilePercent Then
        m_lTilePercence = pc
        ResizeObjects
    End If
End If


End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
m_bMoving = False
UserControl.MousePointer = vbDefault
End Sub

Private Sub UserControl_Resize()
If m_bInitializes = True Then ResizeObjects
End Sub
