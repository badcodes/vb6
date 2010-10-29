Attribute VB_Name = "mdlFunctions"
'*********************************************************************************************
'
' Shell Extensions - Context Menu Handler
'
' Support functions and declarations
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Use at your own risk.
'
' Tested with:
'              * Windows Me / Windows XP
'              * VB6 SP5
'
' History:
'           08/21/1999 - This code was released
'
'*********************************************************************************************
Option Explicit

Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Public Const PAGE_EXECUTE_READWRITE& = &H40&

Public Declare Function lstrcpynA Lib "kernel32" (lpString1 As Any, lpString2 As Any, ByVal MaxLen As Long) As Long

Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Public Const MF_BYPOSITION = &H400&
Public Const MF_SEPARATOR = &H800&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&

'
' Returns a string from a string pointer
'
Public Function StrFromPtrA(ByVal lpszA As Long) As String

    StrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
    lstrcpyA ByVal StrFromPtrA, ByVal lpszA

End Function

'
' Replaces an entry in a object v-table
'
Public Function ReplaceVTableEntry(ByVal oObject As Long, ByVal nEntry As Integer, ByVal pFunc As Long) As Long

    Dim pFuncOld As Long, pVTableHead As Long
    Dim pFuncTmp As Long, lOldProtect As Long
    ' Object pointer contains a pointer to v-table--copy it to temporary
    ' pVTableHead = *oObject;
    MoveMemory pVTableHead, ByVal oObject, 4
    ' Calculate pointer to specified entry
    pFuncTmp = pVTableHead + (nEntry - 1) * 4
    ' Save address of previous method for return
    ' pFuncOld = *pFuncTmp;
    MoveMemory pFuncOld, ByVal pFuncTmp, 4
    ' Ignore if they're already the same
    If pFuncOld <> pFunc Then
        ' Need to change page protection to write to code
        VirtualProtect pFuncTmp, 4, PAGE_EXECUTE_READWRITE, lOldProtect
        ' Write the new function address into the v-table
        MoveMemory ByVal pFuncTmp, pFunc, 4     ' *pFuncTmp = pfunc;
        ' Restore the previous page protection
        VirtualProtect pFuncTmp, 4, lOldProtect, lOldProtect 'Optional
    End If
    'return address of original proc
    ReplaceVTableEntry = pFuncOld

End Function

'
' QueryContextMenu
'
' Adds the menu items to the context menu. This
' function replaces IContextMenu_QueryContextMenu
' because we need to return a value and VB can
' only implement functions that returns HRESULT.
'
' This: A reference to the object
' hMenu: Menu handle of context menu
' indexMenu: index of first menu item
' idCmdFirst: first command ID
' idCmdLast: last command ID
' uFlags: flags
'
Public Function QueryContextMenu(ByVal This As Object, ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As Long) As Long

    Dim ICtxMenu As Handler
    ' Get a reference to the object
    Set ICtxMenu = This
    ' Call the object implementation
    ' of QueryContextMenu
    QueryContextMenu = ICtxMenu.QueryContextMenu(hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags)
    Set ICtxMenu = Nothing

End Function

