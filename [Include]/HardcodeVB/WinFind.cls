VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWindowFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindowsHelper

Public Class As String
Public Title As String

Private Function IWindowsHelper_DoWindow(ByVal iLevel As Integer, ByVal hWnd As Long) As Long
    
    ' Ignore iLevel
    Dim sClass As String, sTitle As String
    sClass = ClassNameFromWnd(hWnd)
    sTitle = WindowTextFromWnd(hWnd)
    If Title = sEmpty Then
        If Class = sEmpty Then
            ' Can't both be empty
            BugAssert True
        Else
            ' Only class needs to match
            If sClass Like Class Then
                IWindowsHelper_DoWindow = hWnd
            End If
        End If
    Else
        If Class = sEmpty Then
            ' Only title needs to match
            If sTitle Like Title Then
                IWindowsHelper_DoWindow = hWnd
            End If
        Else
            ' Both must match
            If sTitle Like Title Then
                If sClass Like Class Then
                    IWindowsHelper_DoWindow = hWnd
               End If
            End If
        End If
    End If
    
End Function
