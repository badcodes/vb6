VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorNotify
    errFirst = 100
    errInvalidDirectory = 100
    errInvalidType = 101
    errInvalidArgument = 102
    errTooManyNotifications = 103
    errNotificationNotFound = 104
    errNoFileChangeMethod = 105
    errLast = 105
End Enum

Private Sub Class_Initialize()
    BugMessage "Initialize object"
    If fLooping = False Then
        ' Start the wait loop
        Call SetTimer(hNull, 0, 200, AddressOf WaitForNotify)
        BugMessage "Started Timer"
        fLooping = True
    End If
    ' Count this object
    cObject = cObject + 1
End Sub

Private Sub Class_Terminate()
    BugMessage "Terminate object"
    ' Uncount this object
    cObject = cObject - 1
    ' Main loop tests -1 because count is 0 to start
    If cObject = 0 Then cObject = -1
End Sub

Function Connect(notifier As IFileNotifier, sDir As String, _
                 efn As EFILE_NOTIFY, fSubTree As Boolean) As Long
    Connect = hInvalid ' Assume fail
    Dim i As Long, h As Long
    ' Find blank handle space
    For i = 0 To cLastNotify
        If ahNotify(i) = hInvalid Then
            ' Set up notification
            h = FindFirstChangeNotification(sDir, fSubTree, efn)
            Connect = h
            If h = hInvalid Then
                ' Change notification unsupported on remote disks
                If Err.LastDllError <> ERROR_NOT_SUPPORTED Then
                    RaiseError errInvalidArgument
                End If
                Exit Function
            End If
            ' Store information
            ahNotify(i) = h
            With aconNotify(i)
                Set .notifier = notifier
                .sDir = sDir
                .efn = efn
                .fSubTree = fSubTree
            End With
            BugMessage "Connecting object: " & i
            Exit Function
        End If
    Next
    RaiseError errTooManyNotifications
End Function

Function Disconnect(hNotify As Long) As Boolean
    Dim i As Long, f As Boolean
    For i = 0 To cLastNotify
        If ahNotify(i) = hNotify Then
            ' Destroy notification
            f = FindCloseChangeNotification(hNotify)
            Disconnect = True
            BugMessage "Disconnecting object: " & i
            ' Compact the arrays
            Do While i <= cLastNotify
                If ahNotify(i) = hInvalid Then Exit Do
                ahNotify(i) = ahNotify(i + 1)
                With aconNotify(i)
                    Set .notifier = aconNotify(i + 1).notifier
                    .sDir = aconNotify(i + 1).sDir
                    .efn = aconNotify(i + 1).efn
                    .fSubTree = aconNotify(i + 1).fSubTree
                End With
                i = i + 1
            Loop
            Exit Function
        End If
    Next
    RaiseError errNotificationNotFound
End Function
'
