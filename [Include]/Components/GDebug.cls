VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GDebug"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#If iVBVer <= 5 Then
' This seems to have been left out of VB5, although it is documented
Public Enum LogModeConstant
    vbLogAuto
    vbLogOff
    vbLogToFile
    vbLogToNT
    vbLogOverwrite = &H10
    vbLogThreadID = &H20
End Enum
#End If

' Output flags determine output destination of BugAsserts and messages
#Const afLogfile = 1
#Const afMsgBox = 2
#Const afDebugWin = 4
#Const afAppLog = 8         ' Log to file
#Const afAppLogNT = 16      ' NT event log
#Const afAppLogMask = 8 Or 16

Private secFreq As Currency
#If afDebug And afLogfile Then
Private iLogFile As Integer
#End If
#If afDebug And (afAppLog Or afAppLogNT) Then
Private fAppLog As Boolean
#End If

Function BugInit() As Boolean
   ' If secFreq = 0 Then BugInit = QueryPerformanceCounter(secFreq)
End Function

Sub BugTerm()
#If afDebug And afLogfile Then
    ' Close log file
    Close iLogFile
    iLogFile = 0
#End If
End Sub

' Display appropriate error message, and then stop
' program.  These errors should NOT be possible in
' shipping product.
Sub BugAssert(ByVal fExpression As Boolean, _
              Optional sExpression As String)
#If afDebug Then
    If fExpression Then Exit Sub
    BugMessage "BugAssert failed: " & sExpression
    Stop
#End If
End Sub
    
    
Sub BugMessage(sMsg As String)
#If afDebug And afLogfile Then
    If iLogFile = 0 Then
        iLogFile = FreeFile
        ' Warning: multiple instances can overwrite log file
        Open App.EXEName & ".DBG" For Output Shared As iLogFile
        ' Challenge: Rewrite to give each instance its own log file
    End If
    Print #iLogFile, sMsg
#End If
#If afDebug And afMsgBox Then
    MsgBox sMsg
#End If
#If afDebug And afDebugWin Then
    Debug.Print sMsg
#End If
#If afDebug And afAppLogMask Then
    If fAppLog = False Then
        fAppLog = True
#If (afDebug And afAppLogMask) = afAppLogNT Then
        App.StartLogging App.Path & "\" & App.EXEName & ".LOG", _
                         vbLogToNT
#ElseIf (afDebug And afAppLogMask) = afAppLog Then
        App.StartLogging App.Path & "\" & App.EXEName & ".LOG", _
                         vbLogToFile Or vbLogOverwrite
#Else
        App.StartLogging App.Path & "\" & App.EXEName & ".LOG", _
                         vbLogAuto Or vbLogOverwrite
#End If
    End If
    App.LogEvent sMsg
#End If
End Sub

Sub BugLocalMessage(sMsg As String)
#If fDebugLocal Then
    BugMessage sMsg
#End If
End Sub

'Sub ProfileStart(secStart As Currency)
'    If secFreq = 0 Then QueryPerformanceFrequency secFreq
'    QueryPerformanceCounter secStart
'End Sub
'
'Sub ProfileStop(secStart As Currency, secTiming As Currency)
'    QueryPerformanceCounter secTiming
'    If secFreq = 0 Then
'        secTiming = 0 ' Handle no high-resolution timer
'    Else
'        secTiming = (secTiming - secStart) / secFreq
'    End If
'End Sub

'Sub ProfileStopMessage(sOutput As String, sPrefix As String, _
'                       secStart As Currency, sPost As String)
'#If afDebug Then
'    Static secTiming As Currency
'    QueryPerformanceCounter secTiming
'    If secFreq = 0 Then
'        secTiming = 0 ' Handle no high-resolution timer
'    Else
'        secTiming = (secTiming - secStart) / secFreq
'    End If
'    ' Return through parameter so that routine can be Sub
'    sOutput = sPrefix & secTiming & sPost
'#End If
'End Sub
'
'Sub BugProfileStop(sPrefix As String, secStart As Currency)
'#If afDebug Then
'    Static secTiming As Currency
'    QueryPerformanceCounter secTiming
'    If secFreq = 0 Then
'        secTiming = 0 ' Handle no high-resolution timer
'    Else
'        secTiming = secTiming - secStart / secFreq
'    End If
'    BugMessage sPrefix & secTiming & " sec "
'#End If
'End Sub

Public Sub PrintError(ByVal vApp As String, ByVal vText As String)
    Debug.Print "Error:" & vApp & ":" & vText
End Sub



