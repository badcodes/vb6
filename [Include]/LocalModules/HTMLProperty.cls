VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHtmlProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mFileName As String
Private Type pmHtmlTag
    name As String
    text As String
'    contentBeginPos As Long
'    contentLen As Long
End Type
Private mHtmlTags() As pmHtmlTag 'start at 1
Private tagCount As Long
Private mLoaded As Boolean
'property stored like this:
' "<a><href>"123"<alt>"abc">"
Public Property Get filename() As String
filename = mFileName
End Property
Public Function GetPropertys(ByRef propArr() As String, ByVal propName As String, Optional ByVal tagName As String) As Long
Attribute GetPropertys.VB_Description = "Store propertys in array start at 1 And return upBound of the array."
'<EhHeader>
On Error GoTo GetPropertys_Err
'</EhHeader>
If mLoaded = False Then Exit Function
Dim i As Long
Dim tagText As String
Dim propValue As String
For i = 1 To tagCount
    If tagName = "" Or StrComp(tagName, mHtmlTags(i).name, vbTextCompare) = 0 Then
        tagText = mHtmlTags(i).text
        If tagText = "" Then GoTo forContinue
        propValue = linvblib.LeftRange(tagText, "<" & propName & ">", "<", vbTextCompare, ReturnEmptyStr)
        If propValue = "" Then GoTo forContinue
        GetPropertys = GetPropertys + 1
        ReDim Preserve propArr(1 To GetPropertys) As String
        propArr(GetPropertys) = propValue
    End If
forContinue:
Next
'<EhFooter>
Exit Function

GetPropertys_Err:
    Err.Raise vbObjectError + 100, _
              "htmlParser.CHtmLDocument.GetPropertys", _
              "CHtmLDocument component failure" & _
              vbCrLf & Err.Description
'</EhFooter>
End Function
'Private Function getWellText() As String
''<EhHeader>
'On Error GoTo getWellText_Err
''</EhHeader>
'Dim fNum As Integer
'Dim wellText As New LiNVBLib.CAppendString
'Dim c As String
'fNum = FreeFile
'Open mFileName For Input As #fNum
'
'Do While Not EOF(fNum)
'    c = strUntil("<", fNum)
'    wellText.Append c & "<"
'    If EOF(fNum) Then Exit Do
'    c = skipChar(" ", fNum)
'    wellText.Append c
'    If c = "/" Then
'        c = skipChar(" ", fNum)
'        wellText.Append c
'    End If
'Loop
'getWellText = wellText.Value
'Close fNum
''<EhFooter>
'Exit Function
'
'getWellText_Err:
'    Err.Raise vbObjectError + 100, _
'              "htmlParser.CHtmLDocument.getWellText", _
'              "CHtmLDocument component failure" & _
'              vbCrLf & Err.Description
''</EhFooter>
'End Function


Private Function getTag(ByRef fNum As Integer) As String
    Dim c As String
    c = skipUntil("<", fNum)
    c = Read(fNum)
    'If skipChar(" ", fNum) Then c = Read(fNum)
    getTag = StrUntil("> ", fNum)
End Function
Private Function getPropertyText(ByRef fNum As Integer) As String

    Dim c As String
    Dim strHandle As New linvblib.CAppendString
    Dim propName As String
    Dim propValue As String
    Dim endChar As String
    
    Do While c <> ">" And (Not EOF(fNum))
        propName = ""
        propValue = ""
        If skipChar(" ", fNum) = False Then GoTo endOfGetProperty
        propName = StrUntil("> =", fNum)
        If propName <> "" Then strHandle.Append "<" & propName & ">"
        
        If skipChar(" ", fNum) Then c = Read(fNum)
        If c = ">" Then GoTo endOfGetProperty
        If c <> "=" Then UnRead (fNum): GoTo loopGetProperty
        
        If skipChar(" ", fNum) Then
            c = Read(fNum)
        Else
            GoTo endOfGetProperty
        End If
        
        If c = Chr(34) Then
            endChar = Chr(34)
        Else
            propValue = c
            endChar = " "
        End If
        
        propValue = propValue & StrUntil(endChar & ">", fNum)
        If propValue <> "" Then strHandle.Append propValue
        c = Read(fNum)
        
loopGetProperty:
Loop
endOfGetProperty:
strHandle.Append "<"
getPropertyText = strHandle.Value
End Function

Public Function InitFrom(ByVal srcPath As String) As Boolean
Attribute InitFrom.VB_Description = "Initialize class from a local htmlfile,return TRUE if success."
'<EhHeader>
On Error GoTo InitFrom_Err
'</EhHeader>
mLoaded = False
InitFrom = False
If linvblib.gCFileSystem.FileExists(srcPath) = False Then Exit Function
mFileName = srcPath
Erase mHtmlTags
tagCount = 0

Dim fNum As Integer
fNum = FreeFile()
Open srcPath For Input Access Read Shared As fNum
Call getall(fNum)
Close (fNum)
InitFrom = True
mLoaded = True
'<EhFooter>
Exit Function

InitFrom_Err:
    Err.Raise vbObjectError + 100, _
              "htmlParser.CHtmLDocument.InitFrom", _
              "CHtmLDocument component failure" & _
              vbCrLf & Err.Description
'</EhFooter>
End Function
Private Sub getall(fNum As Integer)
Dim tmpTag As String
    Do While Not EOF(fNum)
        tmpTag = getTag(fNum)
        If tmpTag <> "" Then
            tagCount = tagCount + 1
            ReDim Preserve mHtmlTags(1 To tagCount) As pmHtmlTag
            mHtmlTags(tagCount).name = tmpTag
            mHtmlTags(tagCount).text = getPropertyText(fNum)
        End If
    Loop
End Sub


