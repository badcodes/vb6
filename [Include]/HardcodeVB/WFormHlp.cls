VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWindowToForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindowsHelper

Private nodeinfo As CStack

Public ShowInvisible As Boolean
Public TreeViewControl As TreeView

Private Function IWindowsHelper_DoWindow(ByVal iLevel As Integer, _
                                         ByVal hWnd As Long) As Long
    BugAssert hWnd <> hNull
    If iLevel < 0 Then Exit Function ' Ignore desktop window
    ' Add window title; if none, add bracketed window class
    Static s As String, i As Integer
    s = WindowTextLineFromWnd(hWnd)
    If s = sEmpty Then s = "[" & ClassNameFromWnd(hWnd) & "]"
    If ShowInvisible Or IsWindowVisible(hWnd) Then
        ' Update treeview control
        With TreeViewControl
            Dim nodX As Node, sKey As String
            ' Create node key
            sKey = "W" & hWnd
            ' Check if this is first node
            If .Nodes.Count = 0 Then
                ' Add node
                Set nodX = .Nodes.Add(, , sKey, s)
                ' Save node info
                Set nodeinfo = New CStack
                nodeinfo.Push CVar(Array(iLevel, sKey))
            Else
                Dim vNode As Variant, tvw As TreeRelationshipConstants
                ' Find parent or sibling
                Do
                    vNode = nodeinfo.Pop
                Loop While (vNode(0) > iLevel)

                ' Add node as sibling or child
                tvw = IIf(vNode(0) = iLevel, tvwNext, tvwChild)
                Set nodX = .Nodes.Add(vNode(1), tvw, sKey, s)
                ' If current node is a child, push parent back on stack
                If vNode(0) < iLevel Then nodeinfo.Push vNode
                ' Push node info on stack
                nodeinfo.Push CVar(Array(iLevel, sKey))
            End If
        End With
    End If
    ' Always successful
    IWindowsHelper_DoWindow = hNull
End Function
'
