VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTemplateBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements IFilterReporter


Private mSource As String
Private mTarget As String
Private mType As CType
Private mFilters() As ITemplateFilter
Private cFilters As Long
Private mReportor As IFilterReporter

Public Function Process(ByVal sSrcFile As String, ByVal sTargetFile As String) As Boolean
    On Error GoTo Process_Abort
    Dim sLines() As String
    Dim cLine As Long
    Dim nOutput As Integer
    Dim nInput As Integer
    nInput = FreeFile
    Open sSrcFile For Input As nInput
    Do While Not EOF(nInput)
        ReDim Preserve sLines(0 To cLine) As String
        Line Input #nInput, sLines(cLine)
        cLine = cLine + 1
    Loop
    Close nInput
    
    Debug.Print "Start Processing " & sSrcFile
    Dim fModified As Boolean
    Dim sTemp As String
    Dim i As Long
    For i = 0 To cFilters - 1
        If (mFilters(i).Process(sLines(), cLine, mType, Me)) Then
            fModified = True
            cLine = SafeUBound(sLines()) - SafeLBound(sLines()) + 1
        End If
    Next
    
    Debug.Print "End Processing " & sSrcFile
    
#If fDebug Then
    'Debug.Print sInput
#Else
    nOutput = FreeFile
    Open sTargetFile For Output As nOutput
    For i = 0 To cLine - 1
        Print #nOutput, sLines(i)
    Next
    Close nOutput
#End If

    Process = True
    Exit Function
Process_Abort:
    Dim nErr As Long
    nErr = Err.Number
    On Error Resume Next
    Close nInput
    Close nOutput
    Err.Raise nErr
End Function

Public Sub AddType(sID As String, sName As String, Optional ttStyle As CTTypeStyles = CTTypeVariant)
    mType.Register sID, sName, ttStyle
End Sub

Public Sub AddFilter(vIFilter As ITemplateFilter)
    ReDim Preserve mFilters(0 To cFilters) As ITemplateFilter
    Set mFilters(cFilters) = vIFilter
    cFilters = cFilters + 1
End Sub

Private Sub Class_Initialize()
    Set mType = New CType
    Set mReportor = Me
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To cFilters - 1
        Set mFilters(i) = Nothing
    Next
    Set mType = Nothing
End Sub

Public Sub InitType(ByRef OType As CType)
    If Not OType Is Nothing Then Set mType = OType
End Sub

Private Sub IFilterReporter_NotifyChange(sFrom As String, sTo As String, Optional nIndex As Long = -1)
    Debug.Print "#" & CStr(nIndex) & ": [" & sFrom & "] To [" & sTo & "]"
End Sub

Private Sub IFilterReporter_RequestChange(nLineIndex As Long, sText As String)
End Sub

Public Property Get Reportor() As IFilterReporter
    Set Reportor = mReportor
End Property

Public Property Set Reportor(ByRef vReportor As IFilterReporter)
    If vReportor Is Nothing Then Exit Property
    Set mReportor = vReportor
End Property

Public Function SafeUBound(vArray As Variant) As Long
    On Error Resume Next
    SafeUBound = -1
    SafeUBound = UBound(vArray)
End Function

Public Function SafeLBound(vArray As Variant) As Long
    On Error Resume Next
    SafeLBound = -1
    SafeLBound = LBound(vArray)
End Function
