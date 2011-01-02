Attribute VB_Name = "MFileNotify"
Option Explicit

Public Type TConnection
    sDir As String
    efn As EFILE_NOTIFY
    fSubTree As Boolean
    notifier As IFileNotifier
End Type

' Actually cLastNotify + 1 allowed
Public Const cLastNotify = 28
' One extra blank item in each array for easy compacting
Public ahNotify(0 To cLastNotify + 1) As Long
Public aconNotify(0 To cLastNotify + 1) As TConnection
Public aerr(errFirst To errLast) As String
' Count of connected objects managed by class
Public cObject As Long
Public fLooping As Boolean

Sub Main()
       
    Dim i As Integer
    For i = 0 To cLastNotify
        ahNotify(i) = hInvalid
    Next
    aerr(errInvalidDirectory) = "Invalid directory"
    aerr(errInvalidType) = "Invalid notification type"
    aerr(errInvalidArgument) = "Invalid argument"
    aerr(errTooManyNotifications) = "Too many notifications"
    aerr(errNotificationNotFound) = "Notification not found"
    BugMessage "Initialized static data"

End Sub

Sub WaitForNotify(ByVal hWnd As Long, ByVal iMsg As Long, _
                  ByVal idTimer As Long, ByVal cCount As Long)
    ' Ignore all parameters except idTimer
    
    ' This one-time callback is used only to start the loop
    KillTimer hNull, idTimer
    BugMessage "Killed Timer"

    Dim iStatus As Long, f As Boolean
    ' Keep waiting for file change events until no more objects
    Do
        '  Wait 100 milliseconds for notification
        iStatus = WaitForMultipleObjects(Count, ahNotify(0), _
                                         False, 100)
        Select Case iStatus
        Case WAIT_TIMEOUT
            ' Nothing happened
            DoEvents
        Case 0 To Count
            BugMessage "Got a notification"
            ' Ignore errors from client; that's their problem
            On Error Resume Next
            ' Call client object with information
            With aconNotify(iStatus)
                .notifier.Change .sDir, .efn, .fSubTree
            End With
            BugMessage "Called back to client"
            ' Wait for next notification
            f = FindNextChangeNotification(ahNotify(iStatus))
        Case WAIT_FAILED
            ' Indicates no notification requests
            DoEvents
        Case Else
            BugMessage "Can't happen"
        End Select
        If cObject < 0 Then BugMessage "Object count: " & cObject
    ' Class Initialize and Terminate events keep reference count
    Loop Until cObject = -1
End Sub

Private Property Get Count() As Long
    Dim i As Long
    For i = 0 To cLastNotify
        If ahNotify(i) = INVALID_HANDLE_VALUE Then Exit For
    Next
    Count = i
End Property

Public Sub RaiseError(iErr As Integer)
    Err.Raise vbObjectError + iErr, "FileNotify.CFileNotify", aerr(iErr)
End Sub
    

