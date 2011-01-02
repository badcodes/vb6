VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "SysTray.ocx"
Begin VB.Form trayform 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin SysTrayCtl.cSysTray cSysTray 
      Left            =   1320
      Top             =   960
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "trayform.frx":0000
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin VB.Menu mnusystem 
      Caption         =   "System"
      Begin VB.Menu menuMain 
         Caption         =   "Main"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "trayform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public thetext As String
Const MyModifiers = MOD_CONTROL + MOD_ALT
Const Mykey = vbKeyL
Const MYTraytip = "Press Ctrl+Shift+L to Lookout"

Private Sub cSysTray_MouseDown(Button As Integer, id As Long)

If Button = 1 Then
    cSysTray.InTray = False
    Unload Me
    Load MainFrm
    MainFrm.Combo2.Text = thetext
    MainFrm.Show
ElseIf Button = 2 Then
    Dim cpos As POINTAPI
    GetCursorPos cpos
    TrackPopupMenu GetSubMenu(GetMenu(Me.hwnd), 0), (TPM_LEFTALIGN Or TPM_RIGHTBUTTON), cpos.x, cpos.y, 0, Me.hwnd, vbNull
End If
    
End Sub


Private Sub Form_Load()

Dim ret As Long
     '记录原来的window程序地址
     preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
     '用自定义程序代替原来的window程序
     ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf wndproc)
     idHotKey = &H1345 'in the range ＆h0000 through ＆hBFFF
     Modifiers = MyModifiers
     uVirtKey = Mykey
     '注册热键
     ret = RegisterHotKey(Me.hwnd, idHotKey, Modifiers, uVirtKey)

With cSysTray
    .InTray = True
    .TrayTip = MYTraytip
End With


 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim ret As Long
     '取消Message的截取，使之送往原来的window程序
     ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, preWinProc)
     Call UnregisterHotKey(Me.hwnd, uVirtKey)

End Sub

Private Sub menuExit_Click()
End
End Sub

Public Sub MenuMain_Click()
    cSysTray.InTray = False
    Unload Me
    Load MainFrm
    MainFrm.Combo2.Text = thetext
    MainFrm.Show
End Sub




