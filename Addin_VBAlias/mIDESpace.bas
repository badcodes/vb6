Attribute VB_Name = "modIDESpace"
Option Explicit

Public MaximiseAtStartup       As Boolean      ' If true, the IDE will start with
' maximised designer and code window.

'LinkedWindows
'don't fight against Madame IDE, just convence her to love you...
Private Type LinkedWnd                        ' To store the references of
    LW                          As Window     ' linked windows which will
End Type                                      ' be hidden.

Public LWArr()                  As LinkedWnd      ' hide-show linked windows
Public k                        As Integer     ' hide-show linked windows

'several operations will look for the user like just one
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'misc. variables
Public LinkedWindowsVisible     As Boolean ' to know if the LnkWnds are visible or hidden
'Public MouseIsDown              As Boolean ' don't process the hook

Public Sub HideLinkedWindows()

'We hide all linked windows and store on the same time the relevant reference
'in an array. Many IDE windows are possibly already hidden and we need to know which
'ones are to be restored later. The point is that when windows are in an
'invisible state, they are also no longer linked.
'The variable k is not really necessary, but I found more confortable coding
'with no use of Ubound. LockWindowUpdate is necessary to provide a nice look.
'In addition, Windows doesn't need to redraw many times the screen.

Dim w As Window
Dim i As Integer

LockWindowUpdate IDEhwnd

With modMacros.VBInstance.ActiveVBProject.VBE
    k = CountLinkedWindows

    If k = 0 Then Exit Sub

    ReDim LWArr(1 To k) 'cleaning and dimensioning the array

    For Each w In .Windows
        If Not w.LinkedWindowFrame Is Nothing Then
            i = i + 1
            Set LWArr(i).LW = w
            w.Visible = False
        End If
    Next w
End With

LockWindowUpdate 0&

LinkedWindowsVisible = False
End Sub
Public Function CountLinkedWindows() As Integer
'my first attempt to use the IDE object. Despite redundant, I left this code
'unchanged as souvenir :)

Dim w As Window, i As Integer

With modSubclass.VBInstance.ActiveVBProject.VBE
    For Each w In .Windows
        If Not w.LinkedWindowFrame Is Nothing Then
            'CountLinkedWindows = w.LinkedWindowFrame.LinkedWindows.Count
            i = i + 1
            'Exit For
        End If
    Next w
        
    CountLinkedWindows = i
End With

End Function
Public Sub ShowLinkedWindows()
'The windows must be restored with inverted placement order, otherwise they
'can't keep the original size. As hidden, they are no longer linked, but the
'linking information is stored in the registry. I will write another little
'but useful program to show how to access and use such informations.

Dim w As Window
Dim i As Integer

If k = 0 Then Exit Sub
LockWindowUpdate IDEhwnd

With modSubclass.VBInstance.ActiveVBProject.VBE
    For i = k To 1 Step -1
        LWArr(i).LW.Visible = True
    Next i
End With

LinkedWindowsVisible = True
LockWindowUpdate 0&
End Sub
Public Sub UnRefLinkedWindows()
'The windows must be restored with inverted placement order, otherwise they
'can't keep the original size. As hidden, they are no longer linked, but the
'linking information is stored in the registry. I will write another little
'but useful program to show how to access and use such informations.

Dim w As Window
Dim i As Integer

If k = 0 Then Exit Sub

With modSubclass.VBInstance.ActiveVBProject.VBE
    For i = k To 1 Step -1
        Set LWArr(i).LW = Nothing
    Next i
End With

ReDim LWArr(0)

End Sub

