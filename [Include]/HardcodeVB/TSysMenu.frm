VERSION 5.00
Begin VB.Form FTestSysMenu 
   Caption         =   "Test System Menu Callback"
   ClientHeight    =   3675
   ClientLeft      =   2355
   ClientTop       =   3480
   ClientWidth     =   4890
   Icon            =   "TSYSMENU.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4890
   Begin VB.Label Label1 
      Caption         =   "Check out About on the system menu. "
      Height          =   525
      Left            =   720
      TabIndex        =   0
      Top             =   690
      Width           =   3435
   End
End
Attribute VB_Name = "FTestSysMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RepId" ,"4FB51841-CEAF-11CF-A15E-00AA00A74D48-0050"
Option Explicit

Private Sub Form_Load()
    Dim hSysMenu As Long
    ' Get handle of system menu
    hSysMenu = GetSystemMenu(hWnd, 0&)
    ' Append separator and menu item with ID IDM_ABOUT
    Call AppendMenu(hSysMenu, MF_SEPARATOR, 0&, 0&)
    Call AppendMenu(hSysMenu, MF_STRING, IDM_ABOUT, "About...")
    Show
    
    ' Install system menu window procedure
    procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SysMenuProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
End Sub

