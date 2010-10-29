Attribute VB_Name = "mdlContextMenu"
Option Explicit

Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long

Public Const MF_BYPOSITION = &H400&
Public Const MF_SEPARATOR = &H800&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Private Function IContextMenu_QueryContextMenu(ByVal This As IContextMenu, ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As olelib.QueryContextMenuFlags) As Long
Dim oCallback As IContextMenuCallback

   Set oCallback = This
   IContextMenu_QueryContextMenu = oCallback.QueryContextMenu(hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags)
   
End Function

Sub ReplaceContextMenu(ByVal CM As olelib.IContextMenu)

   ReplaceVTableEntry ObjPtr(CM), 4, AddressOf IContextMenu_QueryContextMenu

End Sub


