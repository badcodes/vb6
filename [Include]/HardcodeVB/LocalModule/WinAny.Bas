Attribute VB_Name = "MWinFindAny"
Option Explicit

'$ Uses DEBUG.BAS UTILITY.BAS WINTYPE.BAS

#If 1 Then
Private Type TFindAny
    sClass As String
    sTitle As String
    hWndTarget As Long
    fCase As Boolean
End Type

Function FindAnyWindow(Optional Title As String, _
                       Optional Class As String, _
                       Optional CaseSense As Boolean = True) As Long
                       
    ' Pass Title, Class, or both, but not neither
    BugAssert Title <> sEmpty Or Class <> sEmpty
    ' Store parameters in UDT
    Dim find As TFindAny, f As Long
    find.fCase = CaseSense
    find.sClass = IIf(find.fCase, Class, UCase(Class))
    find.sTitle = IIf(find.fCase, Title, UCase(Title))
    ' Ask FindHelper to check each window
    f = EnumChildWindows(GetDesktopWindow, AddressOf FindHelper, find)
    FindAnyWindow = find.hWndTarget

End Function

Private Function FindHelper(ByVal hWnd As Long, find As TFindAny) As Long
    
    Dim sClass As String, sTitle As String
    sClass = MWinTool.ClassNameFromWnd(hWnd)
    sTitle = MWinTool.WindowTextFromWnd(hWnd)
    If find.fCase = False Then
        sTitle = UCase$(sTitle)
        sClass = UCase$(sClass)
    End If
    If find.sTitle = sEmpty Then
        If find.sClass = sEmpty Then
            ' Can't both be empty
            BugAssert True
        Else
            ' Only class needs to match
            If sClass Like find.sClass Then
                find.hWndTarget = hWnd
                Exit Function
            End If
        End If
    Else
        If find.sClass = sEmpty Then
            ' Only title needs to match
            If sTitle Like find.sTitle Then
                find.hWndTarget = hWnd
                Exit Function
            End If
        Else
            ' Both must match
            If sTitle Like find.sTitle Then
                If sClass Like find.sClass Then
                    find.hWndTarget = hWnd
                    Exit Function
               End If
            End If
        End If
    End If
    FindHelper = True
End Function

#Else
Private sClassFind As String, sTitleFind As String, fCase As Boolean

Function FindAnyWindow(Optional Title As String, _
                       Optional Class As String, _
                       Optional CaseSense As Boolean = True) As Long
                       
    Static fInFunc As Long
    If fInFunc Then Exit Function
    fInFunc = True
    
    ' Pass Title, Class, or both, but not neither
    BugAssert Title <> sEmpty Or Class <> sEmpty
    ' Store parameters in module-level variables
    fCase = CaseSense
    sClassFind = IIf(fCase, Class, UCase(Class))
    sTitleFind = IIf(fCase, Title, UCase(Title))
    Dim hWndTarget As Long, f As Long
    ' Ask FindHelper to check each window
    f = EnumChildWindows(GetDesktopWindow, AddressOf FindHelper, hWndTarget)
    FindAnyWindow = hWndTarget
    fInFunc = False

End Function

Private Function FindHelper(ByVal hWnd As Long, lParam As Long) As Long
    
    Dim sClass As String, sTitle As String
    sClass = MWinTool.ClassNameFromWnd(hWnd)
    sTitle = MWinTool.WindowTextFromWnd(hWnd)
    If fCase = False Then
        sTitle = UCase$(sTitle)
        sClass = UCase$(sClass)
    End If
    If sTitleFind = sEmpty Then
        If sClassFind = sEmpty Then
            ' Can't both be empty
            BugAssert True
        Else
            ' Only class needs to match
            If sClass Like sClassFind Then
                lParam = hWnd
                Exit Function
            End If
        End If
    Else
        If sClassFind = sEmpty Then
            ' Only title needs to match
            If sTitle Like sTitleFind Then
                lParam = hWnd
                Exit Function
            End If
        Else
            ' Both must match
            If sTitle Like sTitleFind Then
                If sClass Like sClassFind Then
                    lParam = hWnd
                    Exit Function
               End If
            End If
        End If
    End If
    FindHelper = True
End Function
#End If
'


