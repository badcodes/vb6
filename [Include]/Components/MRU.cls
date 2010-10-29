VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ======================================================================================
' Name:     cMRU
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     1 December 2000
'
' Requires: cRegistry.cls
'
' Copyright ?2000 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' A simple MRU list which persists its state to the
' registry.
'
' ======================================================================================

Private m_iMaxCount As Long
Private m_iCount As Long
Private m_sItems() As String

Public Property Get MaxCount() As Long
   MaxCount = m_iMaxCount
End Property
Public Property Let MaxCount(ByVal lMaxCount As Long)
   m_iMaxCount = lMaxCount
   If m_iCount > m_iMaxCount Then
      m_iCount = m_iMaxCount
   End If
End Property
Public Property Get Count() As Long
   Count = m_iCount
End Property
Public Property Get Item(ByVal nIndex As Long)
   Item = m_sItems(nIndex)
End Property
Public Sub Add(ByVal sItem As String)
Dim i As Long
Dim iIdx As Long
   
   For i = 1 To m_iCount
      If m_sItems(i) = sItem Then
         iIdx = i
         Exit For
      End If
   Next i
   
   If iIdx > 0 Then
      ' swap from 1 -> iIdx-1
      For i = iIdx - 1 To 1 Step -1
         m_sItems(i + 1) = m_sItems(i)
      Next i
   Else
      If m_iCount < m_iMaxCount Then
         m_iCount = m_iCount + 1
         ReDim Preserve m_sItems(1 To m_iCount) As String
      End If
      For i = m_iCount - 1 To 1 Step -1
         m_sItems(i + 1) = m_sItems(i)
      Next i
   End If
   m_sItems(1) = sItem
   
End Sub
Public Function Serialise(ByRef cR As cRegistry) As Boolean
Dim i As Long
   cR.ValueType = REG_DWORD
   cR.ValueKey = "MaxCount"
   cR.Value = m_iMaxCount
   cR.ValueKey = "Count"
   cR.ValueType = REG_SZ
   cR.Value = m_iCount
   For i = 1 To m_iCount
      cR.ValueKey = "Item" & i
      cR.Value = m_sItems(i)
   Next i
End Function
Public Function DeSerialise(ByRef cR As cRegistry) As Boolean
Dim i As Long
   cR.ValueType = REG_DWORD
   cR.ValueKey = "MaxCount"
   m_iMaxCount = cR.Value
   If m_iMaxCount < 1 Then m_iMaxCount = 16
   cR.ValueKey = "Count"
   m_iCount = cR.Value
   If m_iCount > m_iMaxCount Then
      m_iCount = m_iMaxCount
   End If
   If m_iCount > 0 Then
      ReDim m_sItems(1 To m_iCount) As String
      cR.ValueType = REG_SZ
      For i = 1 To m_iCount
         cR.ValueKey = "Item" & i
         m_sItems(i) = cR.Value
      Next i
   Else
      m_iCount = 0
   End If
End Function

Private Sub Class_Initialize()
   m_iMaxCount = 16
End Sub

