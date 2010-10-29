Attribute VB_Name = "Common"


Public Const hwnd_top = 0
Public Const hwnd_bottom = 1
Public Const hwnd_topmost = -1
Public Const hwnd_notopmost = -2
Public Const swp_nosize = &H1
Public Const swp_nomove = &H2
Public Const swp_nozorder = &H4
Public Const swp_noredraw = &H8
Public Const swp_noactivate = &H10
Public Const swp_framechanged = &H20        '  the frame changed: send wm_nccalcsize
Public Const swp_showwindow = &H40
Public Const swp_hidewindow = &H80
Public Const swp_nocopybits = &H100
Public Const swp_noownerzorder = &H200
Public Type POINTAPI
    x As Long
    y As Long
End Type



Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As Any) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long



Function rndarray(numarray() As Integer, minnum As Integer, maxnum As Integer)

minnum = Int(minnum)
maxnum = Int(maxnum)

    If minnum > maxnum Then
    minnum = a
    minnum = maxnum
    maxnum = a
    End If
    
Dim numcount As Integer

numcount = maxnum - minnum + 1

ReDim numarray(numcount) As Integer
ReDim temparray(numcount) As Integer

astart = 1
aend = numcount

For i = 1 To numcount
temparray(i) = i
Next
a = 0

With FrmMain.pbarRndarray '½ø¶ÈÌõ

.Min = 1
If numcount < .Min Then .Max = .Min + 1 Else .Max = numcount

For i = 1 To numcount
Randomize Time
.Value = i
thenum = Int(Rnd(Time) * (aend - astart + 1)) + astart

numarray(i) = temparray(thenum)

    If (thenum - astart) < (aend - thednum) Then
    
    For j = thenum To astart + 1 Step -1
    temparray(j) = temparray(j - 1)
    a = a + 1
    Next
    astart = astart + 1

    Else
    For j = thenum To aend - 1 Step 1
    temparray(j) = temparray(j + 1)
    a = a + 1
    Next
    aend = aend - 1

    End If

Next


End With
    

End Function


Public Function toRGB(vbcolor As Long) As String
colorstr = Hex(vbcolor)
If Len(colorstr) > 6 Then toRGB = colorstr: Exit Function
colorstr = String$(6 - Len(colorstr), "0") + colorstr
toRGB = Right(colorstr, 2) + Mid(colorstr, 3, 2) + Left(colorstr, 2)
End Function
