VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_App As IApp
Private m_bUnloaded As Boolean

Private Sub Form_Load()
    
    Set m_App = New CApp
    m_App.Initialize Me
    Me.Caption = m_App.Title
    
    Call m_App.OnLoad
    
End Sub

Private Sub Form_Terminate()
    If Not m_bUnloaded Then
        m_bUnloaded = True
        m_App.OnUnload
        Set m_App = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_bUnloaded Then
        m_bUnloaded = True
        m_App.OnUnload
        Set m_App = Nothing
    End If
End Sub

