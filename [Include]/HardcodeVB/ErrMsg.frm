VERSION 5.00
Begin VB.Form FErrorMessage 
   Caption         =   "Look Up Error Messages"
   ClientHeight    =   2700
   ClientLeft      =   1260
   ClientTop       =   2268
   ClientWidth     =   4476
   Icon            =   "ErrMsg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FErrorMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2700
   ScaleWidth      =   4476
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   216
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   4092
   End
   Begin VB.TextBox txtError 
      Height          =   495
      Left            =   228
      TabIndex        =   1
      Text            =   "0"
      Top             =   384
      Width           =   1215
   End
   Begin VB.CommandButton cmdLookup 
      Caption         =   "&Lookup"
      Default         =   -1  'True
      Height          =   495
      Left            =   1668
      TabIndex        =   0
      Top             =   384
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Enter error in decimal or hex (Basic, C, or Pascal format):"
      Height          =   216
      Left            =   216
      TabIndex        =   3
      Top             =   48
      Width           =   4092
   End
End
Attribute VB_Name = "FErrorMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLookup_Click()
    Dim iMsgId As Long, ret As Long, sVal As String
    Dim sHex As String, sConst As String
    sVal = txtError
    ' Recognize leading & as a hex specifier without following H
    If Left$(sVal, 1) = "&" And UCase$(Mid$(sVal, 2, 1)) <> "H" Then
        sVal = "&H" & Mid$(sVal, 2)
    ' Recognize leading 0x (C format) as a hex specifier
    ElseIf LCase$(Left$(sVal, 2)) = "0x" Then
        sVal = "&H" & Mid$(sVal, 3)
    ' Recognize leading $ (Pasal format) as a hex specifier
    ElseIf LCase$(Left$(sVal, 1)) = "$" Then
        sVal = "&H" & Mid$(sVal, 2)
    End If
    iMsgId = Val(sVal)
    ' Create the error message
    Dim sNum As String, sMsg As String
    sMsg = String$(256, 0)
    ret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                        FORMAT_MESSAGE_IGNORE_INSERTS, _
                        0&, iMsgId, 0&, sMsg, Len(sMsg), ByVal pNull)
    ' Display it
    sConst = ErrToStr(iMsgId)
    If sConst <> sEmpty Then
        sNum = sConst & " = "
        Clipboard.Clear
        Clipboard.SetText sConst
    Else
        sNum = "Error = "
    End If
    sHex = Right$(String$(4, "0") & Hex$(iMsgId), 4)
    sNum = sNum & iMsgId & " (" & sHex & ")" & vbCrLf
    If ret Then
        txtMessage = sNum & vbCrLf & Left$(sMsg, lstrlen(sMsg))
    Else
        txtMessage = sNum & vbCrLf & "No such error"
    End If
    cmdLookup.SetFocus
    txtError.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Activate()
    txtError.SetFocus
End Sub

Private Sub txtError_GotFocus()
    Debug.Print "Got Focus"
    txtError.SelStart = 0
    txtError.SelLength = 255
End Sub

' Error code definitions for Win32 API errors (from WINERROR.H)

Function ErrToStr(e As Long) As String
    Select Case e
    Case 0
        ErrToStr = "ERROR_SUCCESS"
    Case 1
        ErrToStr = "ERROR_INVALID_FUNCTION"
    Case 2
        ErrToStr = "ERROR_FILE_NOT_FOUND"
    Case 3
        ErrToStr = "ERROR_PATH_NOT_FOUND"
    Case 4
        ErrToStr = "ERROR_TOO_MANY_OPEN_FILES"
    Case 5
        ErrToStr = "ERROR_ACCESS_DENIED"
    Case 6
        ErrToStr = "ERROR_INVALID_HANDLE"
    Case 7
        ErrToStr = "ERROR_ARENA_TRASHED"
    Case 8
        ErrToStr = "ERROR_NOT_ENOUGH_MEMORY"
    Case 9
        ErrToStr = "ERROR_INVALID_BLOCK"
    Case 10
        ErrToStr = "ERROR_BAD_ENVIRONMENT"
    Case 11
        ErrToStr = "ERROR_BAD_FORMAT"
    Case 12
        ErrToStr = "ERROR_INVALID_ACCESS"
    Case 13
        ErrToStr = "ERROR_INVALID_DATA"
    Case 14
        ErrToStr = "ERROR_OUTOFMEMORY"
    Case 15
        ErrToStr = "ERROR_INVALID_DRIVE"
    Case 16
        ErrToStr = "ERROR_CURRENT_DIRECTORY"
    Case 17
        ErrToStr = "ERROR_NOT_SAME_DEVICE"
    Case 18
        ErrToStr = "ERROR_NO_MORE_FILES"
    Case 19
        ErrToStr = "ERROR_WRITE_PROTECT"
    Case 20
        ErrToStr = "ERROR_BAD_UNIT"
    Case 21
        ErrToStr = "ERROR_NOT_READY"
    Case 22
        ErrToStr = "ERROR_BAD_COMMAND"
    Case 23
        ErrToStr = "ERROR_CRC"
    Case 24
        ErrToStr = "ERROR_BAD_LENGTH"
    Case 25
        ErrToStr = "ERROR_SEEK"
    Case 26
        ErrToStr = "ERROR_NOT_DOS_DISK"
    Case 27
        ErrToStr = "ERROR_SECTOR_NOT_FOUND"
    Case 28
        ErrToStr = "ERROR_OUT_OF_PAPER"
    Case 29
        ErrToStr = "ERROR_WRITE_FAULT"
    Case 30
        ErrToStr = "ERROR_READ_FAULT"
    Case 31
        ErrToStr = "ERROR_GEN_FAILURE"
    Case 32
        ErrToStr = "ERROR_SHARING_VIOLATION"
    Case 33
        ErrToStr = "ERROR_LOCK_VIOLATION"
    Case 34
        ErrToStr = "ERROR_WRONG_DISK"
    Case 36
        ErrToStr = "ERROR_SHARING_BUFFER_EXCEEDED"
    Case 38
        ErrToStr = "ERROR_HANDLE_EOF"
    Case 39
        ErrToStr = "ERROR_HANDLE_DISK_FULL"
    Case 50
        ErrToStr = "ERROR_NOT_SUPPORTED"
    Case 51
        ErrToStr = "ERROR_REM_NOT_LIST"
    Case 52
        ErrToStr = "ERROR_DUP_NAME"
    Case 53
        ErrToStr = "ERROR_BAD_NETPATH"
    Case 54
        ErrToStr = "ERROR_NETWORK_BUSY"
    Case 55
        ErrToStr = "ERROR_DEV_NOT_EXIST"
    Case 56
        ErrToStr = "ERROR_TOO_MANY_CMDS"
    Case 57
        ErrToStr = "ERROR_ADAP_HDW_ERR"
    Case 58
        ErrToStr = "ERROR_BAD_NET_RESP"
    Case 59
        ErrToStr = "ERROR_UNEXP_NET_ERR"
    Case 60
        ErrToStr = "ERROR_BAD_REM_ADAP"
    Case 61
        ErrToStr = "ERROR_PRINTQ_FULL"
    Case 62
        ErrToStr = "ERROR_NO_SPOOL_SPACE"
    Case 63
        ErrToStr = "ERROR_PRINT_CANCELLED"
    Case 64
        ErrToStr = "ERROR_NETNAME_DELETED"
    Case 65
        ErrToStr = "ERROR_NETWORK_ACCESS_DENIED"
    Case 66
        ErrToStr = "ERROR_BAD_DEV_TYPE"
    Case 67
        ErrToStr = "ERROR_BAD_NET_NAME"
    Case 68
        ErrToStr = "ERROR_TOO_MANY_NAMES"
    Case 69
        ErrToStr = "ERROR_TOO_MANY_SESS"
    Case 70
        ErrToStr = "ERROR_SHARING_PAUSED"
    Case 71
        ErrToStr = "ERROR_REQ_NOT_ACCEP"
    Case 72
        ErrToStr = "ERROR_REDIR_PAUSED"
    Case 80
        ErrToStr = "ERROR_FILE_EXISTS"
    Case 82
        ErrToStr = "ERROR_CANNOT_MAKE"
    Case 83
        ErrToStr = "ERROR_FAIL_I24"
    Case 84
        ErrToStr = "ERROR_OUT_OF_STRUCTURES"
    Case 85
        ErrToStr = "ERROR_ALREADY_ASSIGNED"
    Case 86
        ErrToStr = "ERROR_INVALID_PASSWORD"
    Case 87
        ErrToStr = "ERROR_INVALID_PARAMETER"
    Case 88
        ErrToStr = "ERROR_NET_WRITE_FAULT"
    Case 89
        ErrToStr = "ERROR_NO_PROC_SLOTS"
    Case 100
        ErrToStr = "ERROR_TOO_MANY_SEMAPHORES"
    Case 101
        ErrToStr = "ERROR_EXCL_SEM_ALREADY_OWNED"
    Case 102
        ErrToStr = "ERROR_SEM_IS_SET"
    Case 103
        ErrToStr = "ERROR_TOO_MANY_SEM_REQUESTS"
    Case 104
        ErrToStr = "ERROR_INVALID_AT_INTERRUPT_TIME"
    Case 105
        ErrToStr = "ERROR_SEM_OWNER_DIED"
    Case 106
        ErrToStr = "ERROR_SEM_USER_LIMIT"
    Case 107
        ErrToStr = "ERROR_DISK_CHANGE"
    Case 108
        ErrToStr = "ERROR_DRIVE_LOCKED"
    Case 109
        ErrToStr = "ERROR_BROKEN_PIPE"
    Case 110
        ErrToStr = "ERROR_OPEN_FAILED"
    Case 111
        ErrToStr = "ERROR_BUFFER_OVERFLOW"
    Case 112
        ErrToStr = "ERROR_DISK_FULL"
    Case 113
        ErrToStr = "ERROR_NO_MORE_SEARCH_HANDLES"
    Case 114
        ErrToStr = "ERROR_INVALID_TARGET_HANDLE"
    Case 117
        ErrToStr = "ERROR_INVALID_CATEGORY"
    Case 118
        ErrToStr = "ERROR_INVALID_VERIFY_SWITCH"
    Case 119
        ErrToStr = "ERROR_BAD_DRIVER_LEVEL"
    Case 120
        ErrToStr = "ERROR_CALL_NOT_IMPLEMENTED"
    Case 121
        ErrToStr = "ERROR_SEM_TIMEOUT"
    Case 122
        ErrToStr = "ERROR_INSUFFICIENT_BUFFER"
    Case 123
        ErrToStr = "ERROR_INVALID_NAME"
    Case 124
        ErrToStr = "ERROR_INVALID_LEVEL"
    Case 125
        ErrToStr = "ERROR_NO_VOLUME_LABEL"
    Case 126
        ErrToStr = "ERROR_MOD_NOT_FOUND"
    Case 127
        ErrToStr = "ERROR_PROC_NOT_FOUND"
    Case 128
        ErrToStr = "ERROR_WAIT_NO_CHILDREN"
    Case 129
        ErrToStr = "ERROR_CHILD_NOT_COMPLETE"
    Case 130
        ErrToStr = "ERROR_DIRECT_ACCESS_HANDLE"
    Case 131
        ErrToStr = "ERROR_NEGATIVE_SEEK"
    Case 132
        ErrToStr = "ERROR_SEEK_ON_DEVICE"
    Case 142
        ErrToStr = "ERROR_BUSY_DRIVE"
    Case 144
        ErrToStr = "ERROR_DIR_NOT_ROOT"
    Case 145
        ErrToStr = "ERROR_DIR_NOT_EMPTY"
    Case 146
        ErrToStr = "ERROR_IS_SUBST_PATH"
    Case 147
        ErrToStr = "ERROR_IS_JOIN_PATH"
    Case 148
        ErrToStr = "ERROR_PATH_BUSY"
    Case 151
        ErrToStr = "ERROR_INVALID_EVENT_COUNT"
    Case 152
        ErrToStr = "ERROR_TOO_MANY_MUXWAITERS"
    Case 153
        ErrToStr = "ERROR_INVALID_LIST_FORMAT"
    Case 154
        ErrToStr = "ERROR_LABEL_TOO_LONG"
    Case 155
        ErrToStr = "ERROR_TOO_MANY_TCBS"
    Case 156
        ErrToStr = "ERROR_SIGNAL_REFUSED"
    Case 157
        ErrToStr = "ERROR_DISCARDED"
    Case 158
        ErrToStr = "ERROR_NOT_LOCKED"
    Case 159
        ErrToStr = "ERROR_BAD_THREADID_ADDR"
    Case 160
        ErrToStr = "ERROR_BAD_ARGUMENTS"
    Case 161
        ErrToStr = "ERROR_BAD_PATHNAME"
    Case 162
        ErrToStr = "ERROR_SIGNAL_PENDING"
    Case 164
        ErrToStr = "ERROR_MAX_THRDS_REACHED"
    Case 167
        ErrToStr = "ERROR_LOCK_FAILED"
    Case 170
        ErrToStr = "ERROR_BUSY"
    Case 173
        ErrToStr = "ERROR_CANCEL_VIOLATION"
    Case 174
        ErrToStr = "ERROR_ATOMIC_LOCKS_NOT_SUPPORTED"
    Case 180
        ErrToStr = "ERROR_INVALID_SEGMENT_NUMBER"

    Case 182
        ErrToStr = "ERROR_INVALID_ORDINAL"
    Case 183
        ErrToStr = "ERROR_ALREADY_EXISTS"

    Case 186
        ErrToStr = "ERROR_INVALID_FLAG_NUMBER"
    Case 187
        ErrToStr = "ERROR_SEM_NOT_FOUND"
    Case 190
        ErrToStr = "ERROR_INVALID_MODULETYPE"
    Case 191
        ErrToStr = "ERROR_INVALID_EXE_SIGNATURE"
    Case 192
        ErrToStr = "ERROR_EXE_MARKED_INVALID"
    Case 193
        ErrToStr = "ERROR_BAD_EXE_FORMAT"
    Case 197
        ErrToStr = "ERROR_IOPL_NOT_ENABLED"
    Case 199
        ErrToStr = "ERROR_AUTODATASEG_EXCEEDS_64k"
    Case 200
        ErrToStr = "ERROR_RING2SEG_MUST_BE_MOVABLE"
    Case 203
        ErrToStr = "ERROR_ENVVAR_NOT_FOUND"
    Case 205
        ErrToStr = "ERROR_NO_SIGNAL_SENT"
    Case 206
        ErrToStr = "ERROR_FILENAME_EXCED_RANGE"
    Case 207
        ErrToStr = "ERROR_RING2_STACK_IN_USE"
    Case 208
        ErrToStr = "ERROR_META_EXPANSION_TOO_LONG"
    Case 209
        ErrToStr = "ERROR_INVALID_SIGNAL_NUMBER"
    Case 210
        ErrToStr = "ERROR_THREAD_1_INACTIVE"
    Case 212
        ErrToStr = "ERROR_LOCKED"
    Case 214
        ErrToStr = "ERROR_TOO_MANY_MODULES"
    Case 215
        ErrToStr = "ERROR_NESTING_NOT_ALLOWED"
    Case 216
        ErrToStr = "ERROR_EXE_MACHINE_TYPE_MISMATCH"
    Case 301
        ErrToStr = "ERROR_INVALID_OPLOCK_PROTOCOL"
    Case 230
        ErrToStr = "ERROR_BAD_PIPE"
    Case 231
        ErrToStr = "ERROR_PIPE_BUSY"
    Case 232
        ErrToStr = "ERROR_NO_DATA"
    Case 233
        ErrToStr = "ERROR_PIPE_NOT_CONNECTED"
    Case 234
        ErrToStr = "ERROR_MORE_DATA"
    Case 240
        ErrToStr = "ERROR_VC_DISCONNECTED"
    Case 259
        ErrToStr = "ERROR_NO_MORE_ITEMS"
    Case 266
        ErrToStr = "ERROR_CANNOT_COPY"
    Case 267
        ErrToStr = "ERROR_DIRECTORY"
    Case 288
        ErrToStr = "ERROR_NOT_OWNER"
    Case 298
        ErrToStr = "ERROR_TOO_MANY_POSTS"
    Case 299
        ErrToStr = "ERROR_PARTIAL_COPY"
    Case 300
        ErrToStr = "ERROR_OPLOCK_NOT_GRANTED"

    Case 317
        ErrToStr = "ERROR_MR_MID_NOT_FOUND"
    Case 487
        ErrToStr = "ERROR_INVALID_ADDRESS"
    Case 534
        ErrToStr = "ERROR_ARITHMETIC_OVERFLOW"
    Case 535
        ErrToStr = "ERROR_PIPE_CONNECTED"
    Case 536
        ErrToStr = "ERROR_PIPE_LISTENING"
    Case 995
        ErrToStr = "ERROR_OPERATION_ABORTED"
    Case 996
        ErrToStr = "ERROR_IO_INCOMPLETE"
    Case 997
        ErrToStr = "ERROR_IO_PENDING"
    Case 998
        ErrToStr = "ERROR_NOACCESS"
    Case 999
        ErrToStr = "ERROR_SWAPERROR"
    Case 1001
        ErrToStr = "ERROR_STACK_OVERFLOW"
    Case 1002
        ErrToStr = "ERROR_INVALID_MESSAGE"
    Case 1003
        ErrToStr = "ERROR_CAN_NOT_COMPLETE"
    Case 1004
        ErrToStr = "ERROR_INVALID_FLAGS"
    Case 1005
        ErrToStr = "ERROR_UNRECOGNIZED_VOLUME"
    Case 1006
        ErrToStr = "ERROR_FILE_INVALID"
    Case 1007
        ErrToStr = "ERROR_FULLSCREEN_MODE"
    Case 1008
        ErrToStr = "ERROR_NO_TOKEN"
    Case 1009
        ErrToStr = "ERROR_BADDB"
    Case 1010
        ErrToStr = "ERROR_BADKEY"
    Case 1011
        ErrToStr = "ERROR_CANTOPEN"
    Case 1012
        ErrToStr = "ERROR_CANTREAD"
    Case 1013
        ErrToStr = "ERROR_CANTWRITE"
    Case 1014
        ErrToStr = "ERROR_REGISTRY_RECOVERED"
    Case 1015
        ErrToStr = "ERROR_REGISTRY_CORRUPT"
    Case 1016
        ErrToStr = "ERROR_REGISTRY_IO_FAILED"
    Case 1017
        ErrToStr = "ERROR_NOT_REGISTRY_FILE"
    Case 1018
        ErrToStr = "ERROR_KEY_DELETED"
    Case 1019
        ErrToStr = "ERROR_NO_LOG_SPACE"
    Case 1020
        ErrToStr = "ERROR_KEY_HAS_CHILDREN"
    Case 1021
        ErrToStr = "ERROR_CHILD_MUST_BE_VOLATILE"
    Case 1022
        ErrToStr = "ERROR_NOTIFY_ENUM_DIR"
    Case 1051
        ErrToStr = "ERROR_DEPENDENT_SERVICES_RUNNING"
    Case 1052
        ErrToStr = "ERROR_INVALID_SERVICE_CONTROL"
    Case 1053
        ErrToStr = "ERROR_SERVICE_REQUEST_TIMEOUT"
    Case 1054
        ErrToStr = "ERROR_SERVICE_NO_THREAD"
    Case 1055
        ErrToStr = "ERROR_SERVICE_DATABASE_LOCKED"
    Case 1056
        ErrToStr = "ERROR_SERVICE_ALREADY_RUNNING"
    Case 1057
        ErrToStr = "ERROR_INVALID_SERVICE_ACCOUNT"
    Case 1058
        ErrToStr = "ERROR_SERVICE_DISABLED"
    Case 1059
        ErrToStr = "ERROR_CIRCULAR_DEPENDENCY"
    Case 1060
        ErrToStr = "ERROR_SERVICE_DOES_NOT_EXIST"
    Case 1061
        ErrToStr = "ERROR_SERVICE_CANNOT_ACCEPT_CTRL"
    Case 1062
        ErrToStr = "ERROR_SERVICE_NOT_ACTIVE"
    Case 1063
        ErrToStr = "ERROR_FAILED_SERVICE_CONTROLLER_CONNECT"
    Case 1064
        ErrToStr = "ERROR_EXCEPTION_IN_SERVICE"
    Case 1065
        ErrToStr = "ERROR_DATABASE_DOES_NOT_EXIST"
    Case 1066
        ErrToStr = "ERROR_SERVICE_SPECIFIC_ERROR"
    Case 1067
        ErrToStr = "ERROR_PROCESS_ABORTED"
    Case 1068
        ErrToStr = "ERROR_SERVICE_DEPENDENCY_FAIL"
    Case 1069
        ErrToStr = "ERROR_SERVICE_LOGON_FAILED"
    Case 1070
        ErrToStr = "ERROR_SERVICE_START_HANG"
    Case 1071
        ErrToStr = "ERROR_INVALID_SERVICE_LOCK"
    Case 1072
        ErrToStr = "ERROR_SERVICE_MARKED_FOR_DELETE"
    Case 1073
        ErrToStr = "ERROR_SERVICE_EXISTS"
    Case 1074
        ErrToStr = "ERROR_ALREADY_RUNNING_LKG"
    Case 1075
        ErrToStr = "ERROR_SERVICE_DEPENDENCY_DELETED"
    Case 1076
        ErrToStr = "ERROR_BOOT_ALREADY_ACCEPTED"
    Case 1077
        ErrToStr = "ERROR_SERVICE_NEVER_STARTED"
    Case 1078
        ErrToStr = "ERROR_DUPLICATE_SERVICE_NAME"
    Case 1100
        ErrToStr = "ERROR_END_OF_MEDIA"
    Case 1101
        ErrToStr = "ERROR_FILEMARK_DETECTED"
    Case 1102
        ErrToStr = "ERROR_BEGINNING_OF_MEDIA"
    Case 1103
        ErrToStr = "ERROR_SETMARK_DETECTED"
    Case 1104
        ErrToStr = "ERROR_NO_DATA_DETECTED"
    Case 1105
        ErrToStr = "ERROR_PARTITION_FAILURE"
    Case 1106
        ErrToStr = "ERROR_INVALID_BLOCK_LENGTH"
    Case 1107
        ErrToStr = "ERROR_DEVICE_NOT_PARTITIONED"
    Case 1108
        ErrToStr = "ERROR_UNABLE_TO_LOCK_MEDIA"
    Case 1109
        ErrToStr = "ERROR_UNABLE_TO_UNLOAD_MEDIA"
    Case 1110
        ErrToStr = "ERROR_MEDIA_CHANGED"
    Case 1111
        ErrToStr = "ERROR_BUS_RESET"
    Case 1112
        ErrToStr = "ERROR_NO_MEDIA_IN_DRIVE"
    Case 1113
        ErrToStr = "ERROR_NO_UNICODE_TRANSLATION"
    Case 1114
        ErrToStr = "ERROR_DLL_INIT_FAILED"
    Case 1115
        ErrToStr = "ERROR_SHUTDOWN_IN_PROGRESS"
    Case 1116
        ErrToStr = "ERROR_NO_SHUTDOWN_IN_PROGRESS"
    Case 1117
        ErrToStr = "ERROR_IO_DEVICE"
    Case 1118
        ErrToStr = "ERROR_SERIAL_NO_DEVICE"
    Case 1119
        ErrToStr = "ERROR_IRQ_BUSY"
    Case 1120
        ErrToStr = "ERROR_MORE_WRITES"
    Case 1121
        ErrToStr = "ERROR_COUNTER_TIMEOUT"
    Case 1129
        ErrToStr = "ERROR_EOM_OVERFLOW"
    Case 1130
        ErrToStr = "ERROR_NOT_ENOUGH_SERVER_MEMORY"
    Case 1131
        ErrToStr = "ERROR_POSSIBLE_DEADLOCK"
    Case 1132
        ErrToStr = "ERROR_MAPPED_ALIGNMENT"
    Case 1140
        ErrToStr = "ERROR_SET_POWER_STATE_VETOED"
    Case 1141
        ErrToStr = "ERROR_SET_POWER_STATE_FAILED"
    Case 1150
        ErrToStr = "ERROR_OLD_WIN_VERSION"
    Case 1151
        ErrToStr = "ERROR_APP_WRONG_OS"
    Case 1152
        ErrToStr = "ERROR_SINGLE_INSTANCE_APP"
    Case 1153
        ErrToStr = "ERROR_RMODE_APP"
    Case 1154
        ErrToStr = "ERROR_INVALID_DLL"
    Case 1155
        ErrToStr = "ERROR_NO_ASSOCIATION"
    Case 1156
        ErrToStr = "ERROR_DDE_FAIL"
    Case 1157
        ErrToStr = "ERROR_DLL_NOT_FOUND"
    Case 1158
        ErrToStr = "ERROR_NO_MORE_USER_HANDLES"

    Case 1159
        ErrToStr = "ERROR_MESSAGE_SYNC_ONLY"

    Case 1160
        ErrToStr = "ERROR_SOURCE_ELEMENT_EMPTY"

    Case 1161
        ErrToStr = "ERROR_DESTINATION_ELEMENT_FULL"

    Case 1162
        ErrToStr = "ERROR_ILLEGAL_ELEMENT_ADDRESS"

    Case 1163
        ErrToStr = "ERROR_MAGAZINE_NOT_PRESENT"

    Case 1164
        ErrToStr = "ERROR_DEVICE_REINITIALIZATION_NEEDED"

    Case 1165
        ErrToStr = "ERROR_DEVICE_REQUIRES_CLEANING"

    Case 1166
        ErrToStr = "ERROR_DEVICE_DOOR_OPEN"

    Case 1167
        ErrToStr = "ERROR_DEVICE_NOT_CONNECTED"

    Case 1168
        ErrToStr = "ERROR_NOT_FOUND"

    Case 1169
        ErrToStr = "ERROR_NO_MATCH"

    Case 1170
        ErrToStr = "ERROR_SET_NOT_FOUND"

    Case 1171
        ErrToStr = "ERROR_POINT_NOT_FOUND"

    Case 1172
        ErrToStr = "ERROR_NO_TRACKING_SERVICE"

    Case 1173
        ErrToStr = "ERROR_NO_VOLUME_ID"

    Case 1200
        ErrToStr = "ERROR_BAD_DEVICE"
    Case 1201
        ErrToStr = "ERROR_CONNECTION_UNAVAIL"
    Case 1202
        ErrToStr = "ERROR_DEVICE_ALREADY_REMEMBERED"
    Case 1203
        ErrToStr = "ERROR_NO_NET_OR_BAD_PATH"
    Case 1204
        ErrToStr = "ERROR_BAD_PROVIDER"
    Case 1205
        ErrToStr = "ERROR_CANNOT_OPEN_PROFILE"
    Case 1206
        ErrToStr = "ERROR_BAD_PROFILE"
    Case 1207
        ErrToStr = "ERROR_NOT_CONTAINER"
    Case 1208
        ErrToStr = "ERROR_EXTENDED_ERROR"
    Case 1209
        ErrToStr = "ERROR_INVALID_GROUPNAME"
    Case 1210
        ErrToStr = "ERROR_INVALID_COMPUTERNAME"
    Case 1211
        ErrToStr = "ERROR_INVALID_EVENTNAME"
    Case 1212
        ErrToStr = "ERROR_INVALID_DOMAINNAME"
    Case 1213
        ErrToStr = "ERROR_INVALID_SERVICENAME"
    Case 1214
        ErrToStr = "ERROR_INVALID_NETNAME"
    Case 1215
        ErrToStr = "ERROR_INVALID_SHARENAME"
    Case 1216
        ErrToStr = "ERROR_INVALID_PASSWORDNAME"
    Case 1217
        ErrToStr = "ERROR_INVALID_MESSAGENAME"
    Case 1218
        ErrToStr = "ERROR_INVALID_MESSAGEDEST"
    Case 1219
        ErrToStr = "ERROR_SESSION_CREDENTIAL_CONFLICT"
    Case 1220
        ErrToStr = "ERROR_REMOTE_SESSION_LIMIT_EXCEEDED"
    Case 1221
        ErrToStr = "ERROR_DUP_DOMAINNAME"
    Case 1222
        ErrToStr = "ERROR_NO_NETWORK"
    Case 1223
        ErrToStr = "ERROR_CANCELLED"
    Case 1224
        ErrToStr = "ERROR_USER_MAPPED_FILE"
    Case 1225
        ErrToStr = "ERROR_CONNECTION_REFUSED"
    Case 1226
        ErrToStr = "ERROR_GRACEFUL_DISCONNECT"
    Case 1227
        ErrToStr = "ERROR_ADDRESS_ALREADY_ASSOCIATED"
    Case 1228
        ErrToStr = "ERROR_ADDRESS_NOT_ASSOCIATED"
    Case 1229
        ErrToStr = "ERROR_CONNECTION_INVALID"
    Case 1230
        ErrToStr = "ERROR_CONNECTION_ACTIVE"
    Case 1231
        ErrToStr = "ERROR_NETWORK_UNREACHABLE"
    Case 1232
        ErrToStr = "ERROR_HOST_UNREACHABLE"
    Case 1233
        ErrToStr = "ERROR_PROTOCOL_UNREACHABLE"
    Case 1234
        ErrToStr = "ERROR_PORT_UNREACHABLE"
    Case 1235
        ErrToStr = "ERROR_REQUEST_ABORTED"
    Case 1236
        ErrToStr = "ERROR_CONNECTION_ABORTED"
    Case 1237
        ErrToStr = "ERROR_RETRY"
    Case 1238
        ErrToStr = "ERROR_CONNECTION_COUNT_LIMIT"
    Case 1239
        ErrToStr = "ERROR_LOGIN_TIME_RESTRICTION"
    Case 1240
        ErrToStr = "ERROR_LOGIN_WKSTA_RESTRICTION"
    Case 1241
        ErrToStr = "ERROR_INCORRECT_ADDRESS"
    Case 1242
        ErrToStr = "ERROR_ALREADY_REGISTERED"
    Case 1243
        ErrToStr = "ERROR_SERVICE_NOT_FOUND"
    Case 1244
        ErrToStr = "ERROR_NOT_AUTHENTICATED"
    Case 1245
        ErrToStr = "ERROR_NOT_LOGGED_ON"
    Case 1246
        ErrToStr = "ERROR_CONTINUE"
    Case 1247
        ErrToStr = "ERROR_ALREADY_INITIALIZED"
    Case 1248
        ErrToStr = "ERROR_NO_MORE_DEVICES"
    Case 1249
        ErrToStr = "ERROR_NO_SUCH_SITE"

    Case 1250
        ErrToStr = "ERROR_DOMAIN_CONTROLLER_EXISTS"

    Case 1251
        ErrToStr = "ERROR_DS_NOT_INSTALLED"

    Case 1311
        ErrToStr = "ERROR_NO_LOGON_SERVERS"

    Case 1300
        ErrToStr = "ERROR_NOT_ALL_ASSIGNED"

    Case 1301
        ErrToStr = "ERROR_SOME_NOT_MAPPED"

    Case 1302
        ErrToStr = "ERROR_NO_QUOTAS_FOR_ACCOUNT"

    Case 1303
        ErrToStr = "ERROR_LOCAL_USER_SESSION_KEY"

    Case 1304
        ErrToStr = "ERROR_NULL_LM_PASSWORD"

    Case 1305
        ErrToStr = "ERROR_UNKNOWN_REVISION"

    Case 1306
        ErrToStr = "ERROR_REVISION_MISMATCH"

    Case 1307
        ErrToStr = "ERROR_INVALID_OWNER"

    Case 1308
        ErrToStr = "ERROR_INVALID_PRIMARY_GROUP"

    Case 1309
        ErrToStr = "ERROR_NO_IMPERSONATION_TOKEN"

    Case 1310
        ErrToStr = "ERROR_CANT_DISABLE_MANDATORY"

    Case 1311
            ErrToStr = "ERROR_NO_LOGON_SERVERS                  "

    Case 1312
        ErrToStr = "ERROR_NO_SUCH_LOGON_SESSION"

    Case 1313
        ErrToStr = "ERROR_NO_SUCH_PRIVILEGE"

    Case 1314
        ErrToStr = "ERROR_PRIVILEGE_NOT_HELD"

    Case 1315
        ErrToStr = "ERROR_INVALID_ACCOUNT_NAME"

    Case 1316
        ErrToStr = "ERROR_USER_EXISTS"

    Case 1317
        ErrToStr = "ERROR_NO_SUCH_USER"

    Case 1318
        ErrToStr = "ERROR_GROUP_EXISTS"

    Case 1319
        ErrToStr = "ERROR_NO_SUCH_GROUP"

    Case 1320
        ErrToStr = "ERROR_MEMBER_IN_GROUP"

    Case 1321
        ErrToStr = "ERROR_MEMBER_NOT_IN_GROUP"

    Case 1322
        ErrToStr = "ERROR_LAST_ADMIN"

    Case 1323
        ErrToStr = "ERROR_WRONG_PASSWORD"

    Case 1324
        ErrToStr = "ERROR_ILL_FORMED_PASSWORD"

    Case 1325
        ErrToStr = "ERROR_PASSWORD_RESTRICTION"

    Case 1326
        ErrToStr = "ERROR_LOGON_FAILURE"

    Case 1327
        ErrToStr = "ERROR_ACCOUNT_RESTRICTION"

    Case 1328
        ErrToStr = "ERROR_INVALID_LOGON_HOURS"

    Case 1329
        ErrToStr = "ERROR_INVALID_WORKSTATION"

    Case 1330
        ErrToStr = "ERROR_PASSWORD_EXPIRED"

    Case 1331
        ErrToStr = "ERROR_ACCOUNT_DISABLED"

    Case 1332
        ErrToStr = "ERROR_NONE_MAPPED"

    Case 1333
        ErrToStr = "ERROR_TOO_MANY_LUIDS_REQUESTED"

    Case 1334
        ErrToStr = "ERROR_LUIDS_EXHAUSTED"

    Case 1335
        ErrToStr = "ERROR_INVALID_SUB_AUTHORITY"

    Case 1336
        ErrToStr = "ERROR_INVALID_ACL"

    Case 1337
        ErrToStr = "ERROR_INVALID_SID"

    Case 1338
        ErrToStr = "ERROR_INVALID_SECURITY_DESCR"

    Case 1340
        ErrToStr = "ERROR_BAD_INHERITANCE_ACL"

    Case 1341
        ErrToStr = "ERROR_SERVER_DISABLED"

    Case 1342
        ErrToStr = "ERROR_SERVER_NOT_DISABLED"

    Case 1343
        ErrToStr = "ERROR_INVALID_ID_AUTHORITY"

    Case 1344
        ErrToStr = "ERROR_ALLOTTED_SPACE_EXCEEDED"

    Case 1345
        ErrToStr = "ERROR_INVALID_GROUP_ATTRIBUTES"

    Case 1346
        ErrToStr = "ERROR_BAD_IMPERSONATION_LEVEL"

    Case 1347
        ErrToStr = "ERROR_CANT_OPEN_ANONYMOUS"

    Case 1348
        ErrToStr = "ERROR_BAD_VALIDATION_CLASS"

    Case 1349
        ErrToStr = "ERROR_BAD_TOKEN_TYPE"

    Case 1350
        ErrToStr = "ERROR_NO_SECURITY_ON_OBJECT"

    Case 1351
        ErrToStr = "ERROR_CANT_ACCESS_DOMAIN_INFO"

    Case 1352
        ErrToStr = "ERROR_INVALID_SERVER_STATE"

    Case 1353
        ErrToStr = "ERROR_INVALID_DOMAIN_STATE"

    Case 1354
        ErrToStr = "ERROR_INVALID_DOMAIN_ROLE"

    Case 1355
        ErrToStr = "ERROR_NO_SUCH_DOMAIN"

    Case 1356
        ErrToStr = "ERROR_DOMAIN_EXISTS"

    Case 1357
        ErrToStr = "ERROR_DOMAIN_LIMIT_EXCEEDED"

    Case 1358
        ErrToStr = "ERROR_INTERNAL_DB_CORRUPTION"

    Case 1359
        ErrToStr = "ERROR_INTERNAL_ERROR"

    Case 1360
        ErrToStr = "ERROR_GENERIC_NOT_MAPPED"

    Case 1361
        ErrToStr = "ERROR_BAD_DESCRIPTOR_FORMAT"

    Case 1362
        ErrToStr = "ERROR_NOT_LOGON_PROCESS"

    Case 1363
        ErrToStr = "ERROR_LOGON_SESSION_EXISTS"

    Case 1364
        ErrToStr = "ERROR_NO_SUCH_PACKAGE"

    Case 1365
        ErrToStr = "ERROR_BAD_LOGON_SESSION_STATE"

    Case 1366
        ErrToStr = "ERROR_LOGON_SESSION_COLLISION"

    Case 1367
        ErrToStr = "ERROR_INVALID_LOGON_TYPE"

    Case 1368
        ErrToStr = "ERROR_CANNOT_IMPERSONATE"

    Case 1369
        ErrToStr = "ERROR_RXACT_INVALID_STATE"

    Case 1370
        ErrToStr = "ERROR_RXACT_COMMIT_FAILURE"

    Case 1371
        ErrToStr = "ERROR_SPECIAL_ACCOUNT"

    Case 1372
        ErrToStr = "ERROR_SPECIAL_GROUP"

    Case 1373
        ErrToStr = "ERROR_SPECIAL_USER"

    Case 1374
        ErrToStr = "ERROR_MEMBERS_PRIMARY_GROUP"

    Case 1375
        ErrToStr = "ERROR_TOKEN_ALREADY_IN_USE"

    Case 1376
        ErrToStr = "ERROR_NO_SUCH_ALIAS"

    Case 1377
        ErrToStr = "ERROR_MEMBER_NOT_IN_ALIAS"

    Case 1378
        ErrToStr = "ERROR_MEMBER_IN_ALIAS"

    Case 1379
        ErrToStr = "ERROR_ALIAS_EXISTS"

    Case 1380
        ErrToStr = "ERROR_LOGON_NOT_GRANTED"

    Case 1381
        ErrToStr = "ERROR_TOO_MANY_SECRETS"

    Case 1382
        ErrToStr = "ERROR_SECRET_TOO_LONG"

    Case 1383
        ErrToStr = "ERROR_INTERNAL_DB_ERROR"

    Case 1384
        ErrToStr = "ERROR_TOO_MANY_CONTEXT_IDS"

    Case 1385
        ErrToStr = "ERROR_LOGON_TYPE_NOT_GRANTED"

    Case 1386
        ErrToStr = "ERROR_NT_CROSS_ENCRYPTION_REQUIRED"

    Case 1387
        ErrToStr = "ERROR_NO_SUCH_MEMBER"

    Case 1388
        ErrToStr = "ERROR_INVALID_MEMBER"

    Case 1389
        ErrToStr = "ERROR_TOO_MANY_SIDS"

    Case 1390
        ErrToStr = "ERROR_LM_CROSS_ENCRYPTION_REQUIRED"

    Case 1391
        ErrToStr = "ERROR_NO_INHERITANCE"

    Case 1392
        ErrToStr = "ERROR_FILE_CORRUPT"

    Case 1393
        ErrToStr = "ERROR_DISK_CORRUPT"

    Case 1394
        ErrToStr = "ERROR_NO_USER_SESSION_KEY"

    Case 1395
        ErrToStr = "ERROR_LICENSE_QUOTA_EXCEEDED"

    Case 1400
        ErrToStr = "ERROR_INVALID_WINDOW_HANDLE"
    Case 1401
        ErrToStr = "ERROR_INVALID_MENU_HANDLE"
    Case 1402
        ErrToStr = "ERROR_INVALID_CURSOR_HANDLE"
    Case 1403
        ErrToStr = "ERROR_INVALID_ACCEL_HANDLE"
    Case 1404
        ErrToStr = "ERROR_INVALID_HOOK_HANDLE"
    Case 1405
        ErrToStr = "ERROR_INVALID_DWP_HANDLE"
    Case 1406
        ErrToStr = "ERROR_TLW_WITH_WSCHILD"
    Case 1407
        ErrToStr = "ERROR_CANNOT_FIND_WND_CLASS"
    Case 1408
        ErrToStr = "ERROR_WINDOW_OF_OTHER_THREAD"
    Case 1409
        ErrToStr = "ERROR_HOTKEY_ALREADY_REGISTERED"
    Case 1410
        ErrToStr = "ERROR_CLASS_ALREADY_EXISTS"
    Case 1411
        ErrToStr = "ERROR_CLASS_DOES_NOT_EXIST"
    Case 1412
        ErrToStr = "ERROR_CLASS_HAS_WINDOWS"
    Case 1413
        ErrToStr = "ERROR_INVALID_INDEX"
    Case 1414
        ErrToStr = "ERROR_INVALID_ICON_HANDLE"
    Case 1415
        ErrToStr = "ERROR_PRIVATE_DIALOG_INDEX"
    Case 1416
        ErrToStr = "ERROR_LISTBOX_ID_NOT_FOUND"
    Case 1417
        ErrToStr = "ERROR_NO_WILDCARD_CHARACTERS"
    Case 1418
        ErrToStr = "ERROR_CLIPBOARD_NOT_OPEN"
    Case 1419
        ErrToStr = "ERROR_HOTKEY_NOT_REGISTERED"
    Case 1420
        ErrToStr = "ERROR_WINDOW_NOT_DIALOG"
    Case 1421
        ErrToStr = "ERROR_CONTROL_ID_NOT_FOUND"
    Case 1422
        ErrToStr = "ERROR_INVALID_COMBOBOX_MESSAGE"
    Case 1423
        ErrToStr = "ERROR_WINDOW_NOT_COMBOBOX"
    Case 1424
        ErrToStr = "ERROR_INVALID_EDIT_HEIGHT"
    Case 1425
        ErrToStr = "ERROR_DC_NOT_FOUND"
    Case 1426
        ErrToStr = "ERROR_INVALID_HOOK_FILTER"
    Case 1427
        ErrToStr = "ERROR_INVALID_FILTER_PROC"
    Case 1428
        ErrToStr = "ERROR_HOOK_NEEDS_HMOD"
    Case 1429
        ErrToStr = "ERROR_GLOBAL_ONLY_HOOK"
    Case 1430
        ErrToStr = "ERROR_JOURNAL_HOOK_SET"
    Case 1431
        ErrToStr = "ERROR_HOOK_NOT_INSTALLED"
    Case 1432
        ErrToStr = "ERROR_INVALID_LB_MESSAGE"
    Case 1433
        ErrToStr = "ERROR_SETCOUNT_ON_BAD_LB"
    Case 1434
        ErrToStr = "ERROR_LB_WITHOUT_TABSTOPS"
    Case 1435
        ErrToStr = "ERROR_DESTROY_OBJECT_OF_OTHER_THREAD"
    Case 1436
        ErrToStr = "ERROR_CHILD_WINDOW_MENU"
    Case 1437
        ErrToStr = "ERROR_NO_SYSTEM_MENU"
    Case 1438
        ErrToStr = "ERROR_INVALID_MSGBOX_STYLE"
    Case 1439
        ErrToStr = "ERROR_INVALID_SPI_VALUE"
    Case 1440
        ErrToStr = "ERROR_SCREEN_ALREADY_LOCKED"
    Case 1441
        ErrToStr = "ERROR_HWNDS_HAVE_DIFF_PARENT"
    Case 1442
        ErrToStr = "ERROR_NOT_CHILD_WINDOW"
    Case 1443
        ErrToStr = "ERROR_INVALID_GW_COMMAND"
    Case 1444
        ErrToStr = "ERROR_INVALID_THREAD_ID"
    Case 1445
        ErrToStr = "ERROR_NON_MDICHILD_WINDOW"
    Case 1446
        ErrToStr = "ERROR_POPUP_ALREADY_ACTIVE"
    Case 1447
        ErrToStr = "ERROR_NO_SCROLLBARS"
    Case 1448
        ErrToStr = "ERROR_INVALID_SCROLLBAR_RANGE"
    Case 1449
        ErrToStr = "ERROR_INVALID_SHOWWIN_COMMAND"

    Case 1450
        ErrToStr = "ERROR_NO_SYSTEM_RESOURCES"

    Case 1451
        ErrToStr = "ERROR_NONPAGED_SYSTEM_RESOURCES"

    Case 1452
        ErrToStr = "ERROR_PAGED_SYSTEM_RESOURCES"

    Case 1453
        ErrToStr = "ERROR_WORKING_SET_QUOTA"

    Case 1454
        ErrToStr = "ERROR_PAGEFILE_QUOTA"

    Case 1455
        ErrToStr = "ERROR_COMMITMENT_LIMIT"

    Case 1456
        ErrToStr = "ERROR_MENU_ITEM_NOT_FOUND"

    Case 1457
        ErrToStr = "ERROR_INVALID_KEYBOARD_HANDLE"

    Case 1458
        ErrToStr = "ERROR_HOOK_TYPE_NOT_ALLOWED"

    Case 1459
        ErrToStr = "ERROR_REQUIRES_INTERACTIVE_WINDOWSTATION"

    Case 1460
        ErrToStr = "ERROR_TIMEOUT"

    Case 1461
        ErrToStr = "ERROR_INVALID_MONITOR_HANDLE"

    Case 1500
        ErrToStr = "ERROR_EVENTLOG_FILE_CORRUPT"

    Case 1501
        ErrToStr = "ERROR_EVENTLOG_CANT_START"

    Case 1502
        ErrToStr = "ERROR_LOG_FILE_FULL"

    Case 1503
        ErrToStr = "ERROR_EVENTLOG_FILE_CHANGED"

    Case 1601
        ErrToStr = "ERROR_INSTALL_SERVICE"

    Case 1602
        ErrToStr = "ERROR_INSTALL_USEREXIT"

    Case 1603
        ErrToStr = "ERROR_INSTALL_FAILURE"

    Case 1604
        ErrToStr = "ERROR_INSTALL_SUSPEND"

    Case 1605
        ErrToStr = "ERROR_UNKNOWN_PRODUCT"

    Case 1606
        ErrToStr = "ERROR_UNKNOWN_FEATURE"

    Case 1607
        ErrToStr = "ERROR_UNKNOWN_COMPONENT"

    Case 1608
        ErrToStr = "ERROR_UNKNOWN_PROPERTY"

    Case 1609
        ErrToStr = "ERROR_INVALID_HANDLE_STATE"

    Case 1610
        ErrToStr = "ERROR_BAD_CONFIGURATION"

    Case 1611
        ErrToStr = "ERROR_INDEX_ABSENT"

    Case 1612
        ErrToStr = "ERROR_INSTALL_SOURCE_ABSENT"

    Case 1613
        ErrToStr = "ERROR_BAD_DATABASE_VERSION"

    Case 1614
        ErrToStr = "ERROR_PRODUCT_UNINSTALLED"

    Case 1615
        ErrToStr = "ERROR_BAD_QUERY_SYNTAX"

    Case 1616
        ErrToStr = "ERROR_INVALID_FIELD"

    Case 1700
        ErrToStr = "RPC_S_INVALID_STRING_BINDING"

    Case 1701
        ErrToStr = "RPC_S_WRONG_KIND_OF_BINDING"

    Case 1702
        ErrToStr = "RPC_S_INVALID_BINDING"

    Case 1703
        ErrToStr = "RPC_S_PROTSEQ_NOT_SUPPORTED"

    Case 1704
        ErrToStr = "RPC_S_INVALID_RPC_PROTSEQ"

    Case 1705
        ErrToStr = "RPC_S_INVALID_STRING_UUID"

    Case 1706
        ErrToStr = "RPC_S_INVALID_ENDPOINT_FORMAT"

    Case 1707
        ErrToStr = "RPC_S_INVALID_NET_ADDR"

    Case 1708
        ErrToStr = "RPC_S_NO_ENDPOINT_FOUND"

    Case 1709
        ErrToStr = "RPC_S_INVALID_TIMEOUT"

    Case 1710
        ErrToStr = "RPC_S_OBJECT_NOT_FOUND"

    Case 1711
        ErrToStr = "RPC_S_ALREADY_REGISTERED"

    Case 1712
        ErrToStr = "RPC_S_TYPE_ALREADY_REGISTERED"

    Case 1713
        ErrToStr = "RPC_S_ALREADY_LISTENING"

    Case 1714
        ErrToStr = "RPC_S_NO_PROTSEQS_REGISTERED"

    Case 1715
        ErrToStr = "RPC_S_NOT_LISTENING"

    Case 1716
        ErrToStr = "RPC_S_UNKNOWN_MGR_TYPE"

    Case 1717
        ErrToStr = "RPC_S_UNKNOWN_IF"

    Case 1718
        ErrToStr = "RPC_S_NO_BINDINGS"

    Case 1719
        ErrToStr = "RPC_S_NO_PROTSEQS"

    Case 1720
        ErrToStr = "RPC_S_CANT_CREATE_ENDPOINT"

    Case 1721
        ErrToStr = "RPC_S_OUT_OF_RESOURCES"

    Case 1722
        ErrToStr = "RPC_S_SERVER_UNAVAILABLE"

    Case 1723
        ErrToStr = "RPC_S_SERVER_TOO_BUSY"

    Case 1724
        ErrToStr = "RPC_S_INVALID_NETWORK_OPTIONS"

    Case 1725
        ErrToStr = "RPC_S_NO_CALL_ACTIVE"

    Case 1726
        ErrToStr = "RPC_S_CALL_FAILED"

    Case 1727
        ErrToStr = "RPC_S_CALL_FAILED_DNE"

    Case 1728
        ErrToStr = "RPC_S_PROTOCOL_ERROR"

    Case 1730
        ErrToStr = "RPC_S_UNSUPPORTED_TRANS_SYN"

    Case 1732
        ErrToStr = "RPC_S_UNSUPPORTED_TYPE"

    Case 1733
        ErrToStr = "RPC_S_INVALID_TAG"

    Case 1734
        ErrToStr = "RPC_S_INVALID_BOUND"

    Case 1735
        ErrToStr = "RPC_S_NO_ENTRY_NAME"

    Case 1736
        ErrToStr = "RPC_S_INVALID_NAME_SYNTAX"

    Case 1737
        ErrToStr = "RPC_S_UNSUPPORTED_NAME_SYNTAX"

    Case 1739
        ErrToStr = "RPC_S_UUID_NO_ADDRESS"

    Case 1740
        ErrToStr = "RPC_S_DUPLICATE_ENDPOINT"

    Case 1741
        ErrToStr = "RPC_S_UNKNOWN_AUTHN_TYPE"

    Case 1742
        ErrToStr = "RPC_S_MAX_CALLS_TOO_SMALL"

    Case 1743
        ErrToStr = "RPC_S_STRING_TOO_LONG"

    Case 1744
        ErrToStr = "RPC_S_PROTSEQ_NOT_FOUND"

    Case 1745
        ErrToStr = "RPC_S_PROCNUM_OUT_OF_RANGE"

    Case 1746
        ErrToStr = "RPC_S_BINDING_HAS_NO_AUTH"

    Case 1747
        ErrToStr = "RPC_S_UNKNOWN_AUTHN_SERVICE"

    Case 1748
        ErrToStr = "RPC_S_UNKNOWN_AUTHN_LEVEL"

    Case 1749
        ErrToStr = "RPC_S_INVALID_AUTH_IDENTITY"

    Case 1750
        ErrToStr = "RPC_S_UNKNOWN_AUTHZ_SERVICE"

    Case 1751
        ErrToStr = "EPT_S_INVALID_ENTRY"

    Case 1752
        ErrToStr = "EPT_S_CANT_PERFORM_OP"

    Case 1753
        ErrToStr = "EPT_S_NOT_REGISTERED"

    Case 1754
        ErrToStr = "RPC_S_NOTHING_TO_EXPORT"

    Case 1755
        ErrToStr = "RPC_S_INCOMPLETE_NAME"

    Case 1756
        ErrToStr = "RPC_S_INVALID_VERS_OPTION"

    Case 1757
        ErrToStr = "RPC_S_NO_MORE_MEMBERS"

    Case 1758
        ErrToStr = "RPC_S_NOT_ALL_OBJS_UNEXPORTED"

    Case 1759
        ErrToStr = "RPC_S_INTERFACE_NOT_FOUND"

    Case 1760
        ErrToStr = "RPC_S_ENTRY_ALREADY_EXISTS"

    Case 1761
        ErrToStr = "RPC_S_ENTRY_NOT_FOUND"

    Case 1762
        ErrToStr = "RPC_S_NAME_SERVICE_UNAVAILABLE"

    Case 1763
        ErrToStr = "RPC_S_INVALID_NAF_ID"

    Case 1764
        ErrToStr = "RPC_S_CANNOT_SUPPORT"

    Case 1765
        ErrToStr = "RPC_S_NO_CONTEXT_AVAILABLE"

    Case 1766
        ErrToStr = "RPC_S_INTERNAL_ERROR"

    Case 1767
        ErrToStr = "RPC_S_ZERO_DIVIDE"

    Case 1768
        ErrToStr = "RPC_S_ADDRESS_ERROR"

    Case 1769
        ErrToStr = "RPC_S_FP_DIV_ZERO"

    Case 1770
        ErrToStr = "RPC_S_FP_UNDERFLOW"

    Case 1771
        ErrToStr = "RPC_S_FP_OVERFLOW"

    Case 1772
        ErrToStr = "RPC_X_NO_MORE_ENTRIES"

    Case 1773
        ErrToStr = "RPC_X_SS_CHAR_TRANS_OPEN_FAIL"

    Case 1774
        ErrToStr = "RPC_X_SS_CHAR_TRANS_SHORT_FILE"

    Case 1775
        ErrToStr = "RPC_X_SS_IN_NULL_CONTEXT"

    Case 1777
        ErrToStr = "RPC_X_SS_CONTEXT_DAMAGED"

    Case 1778
        ErrToStr = "RPC_X_SS_HANDLES_MISMATCH"

    Case 1779
        ErrToStr = "RPC_X_SS_CANNOT_GET_CALL_HANDLE"

    Case 1780
        ErrToStr = "RPC_X_NULL_REF_POINTER"

    Case 1781
        ErrToStr = "RPC_X_ENUM_VALUE_OUT_OF_RANGE"

    Case 1782
        ErrToStr = "RPC_X_BYTE_COUNT_TOO_SMALL"

    Case 1783
        ErrToStr = "RPC_X_BAD_STUB_DATA"

    Case 1784
        ErrToStr = "ERROR_INVALID_USER_BUFFER"

    Case 1785
        ErrToStr = "ERROR_UNRECOGNIZED_MEDIA"

    Case 1786
        ErrToStr = "ERROR_NO_TRUST_LSA_SECRET"

    Case 1787
        ErrToStr = "ERROR_NO_TRUST_SAM_ACCOUNT"

    Case 1788
        ErrToStr = "ERROR_TRUSTED_DOMAIN_FAILURE"

    Case 1789
        ErrToStr = "ERROR_TRUSTED_RELATIONSHIP_FAILURE"

    Case 1790
        ErrToStr = "ERROR_TRUST_FAILURE"

    Case 1791
        ErrToStr = "RPC_S_CALL_IN_PROGRESS"

    Case 1792
        ErrToStr = "ERROR_NETLOGON_NOT_STARTED"

    Case 1793
        ErrToStr = "ERROR_ACCOUNT_EXPIRED"

    Case 1794
        ErrToStr = "ERROR_REDIRECTOR_HAS_OPEN_HANDLES"

    Case 1795
        ErrToStr = "ERROR_PRINTER_DRIVER_ALREADY_INSTALLED"

    Case 1796
        ErrToStr = "ERROR_UNKNOWN_PORT"

    Case 1797
        ErrToStr = "ERROR_UNKNOWN_PRINTER_DRIVER"

    Case 1798
        ErrToStr = "ERROR_UNKNOWN_PRINTPROCESSOR"

    Case 1799
        ErrToStr = "ERROR_INVALID_SEPARATOR_FILE"

    Case 1800
        ErrToStr = "ERROR_INVALID_PRIORITY"

    Case 1801
        ErrToStr = "ERROR_INVALID_PRINTER_NAME"

    Case 1802
        ErrToStr = "ERROR_PRINTER_ALREADY_EXISTS"

    Case 1803
        ErrToStr = "ERROR_INVALID_PRINTER_COMMAND"

    Case 1804
        ErrToStr = "ERROR_INVALID_DATATYPE"

    Case 1805
        ErrToStr = "ERROR_INVALID_ENVIRONMENT"

    Case 1806
        ErrToStr = "RPC_S_NO_MORE_BINDINGS"

    Case 1807
        ErrToStr = "ERROR_NOLOGON_INTERDOMAIN_TRUST_ACCOUNT"

    Case 1808
        ErrToStr = "ERROR_NOLOGON_WORKSTATION_TRUST_ACCOUNT"

    Case 1809
        ErrToStr = "ERROR_NOLOGON_SERVER_TRUST_ACCOUNT"

    Case 1810
        ErrToStr = "ERROR_DOMAIN_TRUST_INCONSISTENT"

    Case 1811
        ErrToStr = "ERROR_SERVER_HAS_OPEN_HANDLES"

    Case 1812
        ErrToStr = "ERROR_RESOURCE_DATA_NOT_FOUND"

    Case 1813
        ErrToStr = "ERROR_RESOURCE_TYPE_NOT_FOUND"

    Case 1814
        ErrToStr = "ERROR_RESOURCE_NAME_NOT_FOUND"

    Case 1815
        ErrToStr = "ERROR_RESOURCE_LANG_NOT_FOUND"

    Case 1816
        ErrToStr = "ERROR_NOT_ENOUGH_QUOTA"

    Case 1817
        ErrToStr = "RPC_S_NO_INTERFACES"

    Case 1818
        ErrToStr = "RPC_S_CALL_CANCELLED"

    Case 1819
        ErrToStr = "RPC_S_BINDING_INCOMPLETE"

    Case 1820
        ErrToStr = "RPC_S_COMM_FAILURE"

    Case 1821
        ErrToStr = "RPC_S_UNSUPPORTED_AUTHN_LEVEL"

    Case 1822
        ErrToStr = "RPC_S_NO_PRINC_NAME"

    Case 1823
        ErrToStr = "RPC_S_NOT_RPC_ERROR"

    Case 1824
        ErrToStr = "RPC_S_UUID_LOCAL_ONLY"

    Case 1825
        ErrToStr = "RPC_S_SEC_PKG_ERROR"

    Case 1826
        ErrToStr = "RPC_S_NOT_CANCELLED"

    Case 1827
        ErrToStr = "RPC_X_INVALID_ES_ACTION"

    Case 1828
        ErrToStr = "RPC_X_WRONG_ES_VERSION"

    Case 1829
        ErrToStr = "RPC_X_WRONG_STUB_VERSION"

    Case 1830
        ErrToStr = "RPC_X_INVALID_PIPE_OBJECT"

    Case 1831
        ErrToStr = "RPC_X_WRONG_PIPE_ORDER"

    Case 1832
        ErrToStr = "RPC_X_WRONG_PIPE_VERSION"

    Case 1898
        ErrToStr = "RPC_S_GROUP_MEMBER_NOT_FOUND"

    Case 1899
        ErrToStr = "EPT_S_CANT_CREATE"

    Case 1900
        ErrToStr = "RPC_S_INVALID_OBJECT"

    Case 1901
        ErrToStr = "ERROR_INVALID_TIME"

    Case 1902
        ErrToStr = "ERROR_INVALID_FORM_NAME"

    Case 1903
        ErrToStr = "ERROR_INVALID_FORM_SIZE"

    Case 1904
        ErrToStr = "ERROR_ALREADY_WAITING"

    Case 1905
        ErrToStr = "ERROR_PRINTER_DELETED"

    Case 1906
        ErrToStr = "ERROR_INVALID_PRINTER_STATE"

    Case 1907
        ErrToStr = "ERROR_PASSWORD_MUST_CHANGE"

    Case 1908
        ErrToStr = "ERROR_DOMAIN_CONTROLLER_NOT_FOUND"

    Case 1909
        ErrToStr = "ERROR_ACCOUNT_LOCKED_OUT"

    Case 1910
        ErrToStr = "OR_INVALID_OXID"

    Case 1911
        ErrToStr = "OR_INVALID_OID"

    Case 1912
        ErrToStr = "OR_INVALID_SET"

    Case 1913
        ErrToStr = "RPC_S_SEND_INCOMPLETE"

    Case 1914
        ErrToStr = "RPC_S_INVALID_ASYNC_HANDLE"

    Case 1915
        ErrToStr = "RPC_S_INVALID_ASYNC_CALL"

    Case 1916
        ErrToStr = "RPC_X_PIPE_CLOSED"

    Case 1917
        ErrToStr = "RPC_X_PIPE_DISCIPLINE_ERROR"

    Case 1918
        ErrToStr = "RPC_X_PIPE_EMPTY"

    Case 1919
        ErrToStr = "ERROR_NO_SITENAME"

    Case 1920
        ErrToStr = "ERROR_CANT_ACCESS_FILE"

    Case 1921
        ErrToStr = "ERROR_CANT_RESOLVE_FILENAME"

    Case 1922
        ErrToStr = "ERROR_DS_MEMBERSHIP_EVALUATED_LOCALLY"

    Case 1923
        ErrToStr = "ERROR_DS_NO_ATTRIBUTE_OR_VALUE"

    Case 1924
        ErrToStr = "ERROR_DS_INVALID_ATTRIBUTE_SYNTAX"

    Case 1925
        ErrToStr = "ERROR_DS_ATTRIBUTE_TYPE_UNDEFINED"

    Case 1926
        ErrToStr = "ERROR_DS_ATTRIBUTE_OR_VALUE_EXISTS"

    Case 1927
        ErrToStr = "ERROR_DS_BUSY"

    Case 1928
        ErrToStr = "ERROR_DS_UNAVAILABLE"

    Case 1929
        ErrToStr = "ERROR_DS_NO_RIDS_ALLOCATED"

    Case 1930
        ErrToStr = "ERROR_DS_NO_MORE_RIDS"

    Case 1931
        ErrToStr = "ERROR_DS_INCORRECT_ROLE_OWNER"

    Case 1932
        ErrToStr = "ERROR_DS_RIDMGR_INIT_ERROR"

    Case 1933
        ErrToStr = "ERROR_DS_OBJ_CLASS_VIOLATION"

    Case 1934
        ErrToStr = "ERROR_DS_CANT_ON_NON_LEAF"

    Case 1935
        ErrToStr = "ERROR_DS_CANT_ON_RDN"

    Case 1936
        ErrToStr = "ERROR_DS_CANT_MOD_OBJ_CLASS"

    Case 1937
        ErrToStr = "ERROR_DS_CROSS_DOM_MOVE_ERROR"

    Case 1938
        ErrToStr = "ERROR_DS_GC_NOT_AVAILABLE"

    Case 6118
        ErrToStr = "ERROR_NO_BROWSER_SERVERS_FOUND"

    Case 2000
        ErrToStr = "ERROR_INVALID_PIXEL_FORMAT"

    Case 2001
        ErrToStr = "ERROR_BAD_DRIVER"

    Case 2002
        ErrToStr = "ERROR_INVALID_WINDOW_STYLE"

    Case 2003
        ErrToStr = "ERROR_METAFILE_NOT_SUPPORTED"

    Case 2004
        ErrToStr = "ERROR_TRANSFORM_NOT_SUPPORTED"

    Case 2005
        ErrToStr = "ERROR_CLIPPING_NOT_SUPPORTED"

    Case 2108
        ErrToStr = "ERROR_CONNECTED_OTHER_PASSWORD"

    Case 2202
        ErrToStr = "ERROR_BAD_USERNAME"
    Case 2250
        ErrToStr = "ERROR_NOT_CONNECTED"
    Case 2300
        ErrToStr = "ERROR_INVALID_CMM"

    Case 2301
        ErrToStr = "ERROR_INVALID_PROFILE"

    Case 2302
        ErrToStr = "ERROR_TAG_NOT_FOUND"

    Case 2303
        ErrToStr = "ERROR_TAG_NOT_PRESENT"

    Case 2304
        ErrToStr = "ERROR_DUPLICATE_TAG"

    Case 2305
        ErrToStr = "ERROR_PROFILE_NOT_ASSOCIATED_WITH_DEVICE"

    Case 2306
        ErrToStr = "ERROR_PROFILE_NOT_FOUND"

    Case 2307
        ErrToStr = "ERROR_INVALID_COLORSPACE"

    Case 2308
        ErrToStr = "ERROR_ICM_NOT_ENABLED"

    Case 2309
        ErrToStr = "ERROR_DELETING_ICM_XFORM"

    Case 2310
        ErrToStr = "ERROR_INVALID_TRANSFORM"

    Case 2401
        ErrToStr = "ERROR_OPEN_FILES"
    Case 2402
        ErrToStr = "ERROR_ACTIVE_CONNECTIONS"
    Case 2404
        ErrToStr = "ERROR_DEVICE_IN_USE"
    Case 3000
        ErrToStr = "ERROR_UNKNOWN_PRINT_MONITOR"

    Case 3001
        ErrToStr = "ERROR_PRINTER_DRIVER_IN_USE"

    Case 3002
        ErrToStr = "ERROR_SPOOL_FILE_NOT_FOUND"

    Case 3003
        ErrToStr = "ERROR_SPL_NO_STARTDOC"

    Case 3004
        ErrToStr = "ERROR_SPL_NO_ADDJOB"

    Case 3005
        ErrToStr = "ERROR_PRINT_PROCESSOR_ALREADY_INSTALLED"

    Case 3006
        ErrToStr = "ERROR_PRINT_MONITOR_ALREADY_INSTALLED"

    Case 3007
        ErrToStr = "ERROR_INVALID_PRINT_MONITOR"

    Case 3008
        ErrToStr = "ERROR_PRINT_MONITOR_IN_USE"

    Case 3009
        ErrToStr = "ERROR_PRINTER_HAS_JOBS_QUEUED"

    Case 3010
        ErrToStr = "ERROR_SUCCESS_REBOOT_REQUIRED"

    Case 3011
        ErrToStr = "ERROR_SUCCESS_RESTART_REQUIRED"

    Case 4000
        ErrToStr = "ERROR_WINS_INTERNAL"

    Case 4001
        ErrToStr = "ERROR_CAN_NOT_DEL_LOCAL_WINS"

    Case 4002
        ErrToStr = "ERROR_STATIC_INIT"

    Case 4003
        ErrToStr = "ERROR_INC_BACKUP"

    Case 4004
        ErrToStr = "ERROR_FULL_BACKUP"

    Case 4005
        ErrToStr = "ERROR_REC_NON_EXISTENT"

    Case 4006
        ErrToStr = "ERROR_RPL_NOT_ALLOWED"

    Case 4100
        ErrToStr = "ERROR_DHCP_ADDRESS_CONFLICT"

    Case 4200
        ErrToStr = "ERROR_WMI_GUID_NOT_FOUND"

    Case 4201
        ErrToStr = "ERROR_WMI_INSTANCE_NOT_FOUND"

    Case 4202
        ErrToStr = "ERROR_WMI_ITEMID_NOT_FOUND"

    Case 4203
        ErrToStr = "ERROR_WMI_TRY_AGAIN"

    Case 4204
        ErrToStr = "ERROR_WMI_DP_NOT_FOUND"

    Case 4205
        ErrToStr = "ERROR_WMI_UNRESOLVED_INSTANCE_REF"

    Case 4206
        ErrToStr = "ERROR_WMI_ALREADY_ENABLED"

    Case 4207
        ErrToStr = "ERROR_WMI_GUID_DISCONNECTED"

    Case 4208
        ErrToStr = "ERROR_WMI_SERVER_UNAVAILABLE"

    Case 4209
        ErrToStr = "ERROR_WMI_DP_FAILED"

    Case 4210
        ErrToStr = "ERROR_WMI_INVALID_MOF"

    Case 4211
        ErrToStr = "ERROR_WMI_INVALID_REGINFO"

    Case 4300
        ErrToStr = "ERROR_INVALID_MEDIA"

    Case 4301
        ErrToStr = "ERROR_INVALID_LIBRARY"

    Case 4302
        ErrToStr = "ERROR_INVALID_MEDIA_POOL"

    Case 4303
        ErrToStr = "ERROR_DRIVE_MEDIA_MISMATCH"

    Case 4304
        ErrToStr = "ERROR_MEDIA_OFFLINE"

    Case 4305
        ErrToStr = "ERROR_LIBRARY_OFFLINE"

    Case 4306
        ErrToStr = "ERROR_EMPTY"

    Case 4307
        ErrToStr = "ERROR_NOT_EMPTY"

    Case 4308
        ErrToStr = "ERROR_MEDIA_UNAVAILABLE"

    Case 4309
        ErrToStr = "ERROR_RESOURCE_DISABLED"

    Case 4310
        ErrToStr = "ERROR_INVALID_CLEANER"

    Case 4311
        ErrToStr = "ERROR_UNABLE_TO_CLEAN"

    Case 4312
        ErrToStr = "ERROR_OBJECT_NOT_FOUND"

    Case 4313
        ErrToStr = "ERROR_DATABASE_FAILURE"

    Case 4314
        ErrToStr = "ERROR_DATABASE_FULL"

    Case 4315
        ErrToStr = "ERROR_MEDIA_INCOMPATIBLE"

    Case 4316
        ErrToStr = "ERROR_RESOURCE_NOT_PRESENT"

    Case 4317
        ErrToStr = "ERROR_INVALID_OPERATION"

    Case 4318
        ErrToStr = "ERROR_MEDIA_NOT_AVAILABLE"

    Case 4319
        ErrToStr = "ERROR_DEVICE_NOT_AVAILABLE"

    Case 4320
        ErrToStr = "ERROR_REQUEST_REFUSED"

    Case 4350
        ErrToStr = "ERROR_FILE_OFFLINE"

    Case 4351
        ErrToStr = "ERROR_REMOTE_STORAGE_NOT_ACTIVE"

    Case 4352
        ErrToStr = "ERROR_REMOTE_STORAGE_MEDIA_ERROR"

    Case 4390
        ErrToStr = "ERROR_NOT_A_REPARSE_POINT"

    Case 4391
        ErrToStr = "ERROR_REPARSE_ATTRIBUTE_CONFLICT"

    Case 5001
        ErrToStr = "ERROR_DEPENDENT_RESOURCE_EXISTS"

    Case 5002
        ErrToStr = "ERROR_DEPENDENCY_NOT_FOUND"

    Case 5003
        ErrToStr = "ERROR_DEPENDENCY_ALREADY_EXISTS"

    Case 5004
        ErrToStr = "ERROR_RESOURCE_NOT_ONLINE"

    Case 5005
        ErrToStr = "ERROR_HOST_NODE_NOT_AVAILABLE"

    Case 5006
        ErrToStr = "ERROR_RESOURCE_NOT_AVAILABLE"

    Case 5007
        ErrToStr = "ERROR_RESOURCE_NOT_FOUND"

    Case 5008
        ErrToStr = "ERROR_SHUTDOWN_CLUSTER"

    Case 5009
        ErrToStr = "ERROR_CANT_EVICT_ACTIVE_NODE"

    Case 5010
        ErrToStr = "ERROR_OBJECT_ALREADY_EXISTS"

    Case 5011
        ErrToStr = "ERROR_OBJECT_IN_LIST"

    Case 5012
        ErrToStr = "ERROR_GROUP_NOT_AVAILABLE"

    Case 5013
        ErrToStr = "ERROR_GROUP_NOT_FOUND"

    Case 5014
        ErrToStr = "ERROR_GROUP_NOT_ONLINE"

    Case 5015
        ErrToStr = "ERROR_HOST_NODE_NOT_RESOURCE_OWNER"

    Case 5016
        ErrToStr = "ERROR_HOST_NODE_NOT_GROUP_OWNER"

    Case 5017
        ErrToStr = "ERROR_RESMON_CREATE_FAILED"

    Case 5018
        ErrToStr = "ERROR_RESMON_ONLINE_FAILED"

    Case 5019
        ErrToStr = "ERROR_RESOURCE_ONLINE"

    Case 5020
        ErrToStr = "ERROR_QUORUM_RESOURCE"

    Case 5021
        ErrToStr = "ERROR_NOT_QUORUM_CAPABLE"

    Case 5022
        ErrToStr = "ERROR_CLUSTER_SHUTTING_DOWN"

    Case 5023
        ErrToStr = "ERROR_INVALID_STATE"

    Case 5024
        ErrToStr = "ERROR_RESOURCE_PROPERTIES_STORED"

    Case 5025
        ErrToStr = "ERROR_NOT_QUORUM_CLASS"

    Case 5026
        ErrToStr = "ERROR_CORE_RESOURCE"

    Case 5027
        ErrToStr = "ERROR_QUORUM_RESOURCE_ONLINE_FAILED"

    Case 5028
        ErrToStr = "ERROR_QUORUMLOG_OPEN_FAILED"

    Case 5029
        ErrToStr = "ERROR_CLUSTERLOG_CORRUPT"

    Case 5030
        ErrToStr = "ERROR_CLUSTERLOG_RECORD_EXCEEDS_MAXSIZE"

    Case 5031
        ErrToStr = "ERROR_CLUSTERLOG_EXCEEDS_MAXSIZE"

    Case 5032
        ErrToStr = "ERROR_CLUSTERLOG_CHKPOINT_NOT_FOUND"

    Case 5033
        ErrToStr = "ERROR_CLUSTERLOG_NOT_ENOUGH_SPACE"

    Case 6000
        ErrToStr = "ERROR_ENCRYPTION_FAILED"

    Case 6001
        ErrToStr = "ERROR_DECRYPTION_FAILED"

    Case 6002
        ErrToStr = "ERROR_FILE_ENCRYPTED"

    Case 6003
        ErrToStr = "ERROR_NO_RECOVERY_POLICY"

    Case 6004
        ErrToStr = "ERROR_NO_EFS"

    Case 6005
        ErrToStr = "ERROR_WRONG_EFS"

    Case 6006
        ErrToStr = "ERROR_NO_USER_KEYS"

    Case 6007
        ErrToStr = "ERROR_FILE_NOT_ENCRYPTED"

    Case 6008
        ErrToStr = "ERROR_NOT_EXPORT_FORMAT"

    'Case Else
        'ErrToStr = sEmpty
    End Select
End Function

