Attribute VB_Name = "modtxtReader"

Option Explicit
Type MYPoS
    Top As Long
    Left As Long
    Height As Long
    Width As Long
End Type


Type ReaderStyle
    WindowState As FormWindowStateConstants
    formPos As MYPoS
    LeftWidth As Long
    LastPath As String
    ShowMenu As Boolean
    ShowLeft As Boolean
    ShowStatusBar As Boolean
    FullScreenMode As Boolean
    TextEditor As String
End Type

Type ViewerStyle
    Viewfont As MYFont
    ForeColor As OLE_COLOR
    BackColor As OLE_COLOR
    LineHeight As Integer
    RecentMax As Integer
End Type

Public Type zhReaderStatus
 sPWD As String
 sCur_zhFile As String
 sCur_zhSubFile As String
 bMenuShowed As Boolean
 bLeftShowed As Boolean
 bStatusBarShowed As Boolean
End Type


Public Const TempHtm = "$$TEMP$$.HTM"
Public Const cHtmlAboutFilename = "about.htm"
'Public Const szhHtmlTemplate = "html\book.htm"
Public zhrStatus As zhReaderStatus
Private Const zhMemorySplit = vbCrLf


'Public Const TempHtm = "$$HTML$$.HTM"


Public Sub loadBookmark(iniFiletoDo As String, ByRef mnuBookmark As Object)

    Dim i As Integer
    Dim bCount As Long

        bCount = Val(iniGetSetting(iniFiletoDo, "Bookmark", "Count")) - 1
        On Error Resume Next
        For i = 0 To bCount
            Load mnuBookmark(i + 1)
            mnuBookmark(i + 1).Caption = iniGetSetting(iniFiletoDo, "Bookmark", "Name" & Str$(i))
            mnuBookmark(i + 1).Tag = iniGetSetting(iniFiletoDo, "Bookmark", "Location" & Str$(i))
            mnuBookmark(i + 1).Visible = True
        Next


End Sub

Public Sub saveBookmark(iniFiletoDo As String, ByRef mnuBookmark As Object)

    Dim i As Integer
    iniDeleteSection iniFiletoDo, "Bookmark"
    Dim fNum As Integer
    fNum = FreeFile
    Open iniFiletoDo For Append As fNum
    Print #fNum, "[Bookmark]"
    Print #fNum, "Count=" & mnuBookmark.Count - 1
    For i = 1 To mnuBookmark.Count - 1
        Print #fNum, "Name" & Str$(i - 1) & "=" & mnuBookmark(i).Caption
        Print #fNum, "Location" & Str$(i - 1) & "=" & mnuBookmark(i).Tag
    Next
    Close #fNum
    
End Sub

Public Sub GetReaderStyle(iniFiletoDo As String, RS As ReaderStyle)

    With RS.formPos
        .Height = CLngStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "FormHeight"))
        .Width = CLngStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "FormWidth"))
        .Top = CLngStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "FormTop"))
        .Left = CLngStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "FormLeft"))
    End With

    RS.WindowState = CLngStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "WindowState"))
    RS.LeftWidth = CLngStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "LeftWidth"))
    RS.LastPath = iniGetSetting(iniFiletoDo, "ReaderStyle", "LastPath")
    RS.ShowMenu = CBoolStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "ShowMenu"))
    RS.ShowLeft = CBoolStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "ShowLeft"))
    RS.ShowStatusBar = CBoolStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "ShowStatusBar"))
    RS.FullScreenMode = CBoolStr(iniGetSetting(iniFiletoDo, "ReaderStyle", "FullScreenMode"))
    RS.TextEditor = iniGetSetting(iniFiletoDo, "ReaderStyle", "TextEditor")
End Sub

Public Sub SaveReaderStyle(iniFiletoDo As String, RS As ReaderStyle)

    iniDeleteSection iniFiletoDo, "ReaderStyle"
    Dim fNum As Integer
    fNum = FreeFile
    Open iniFiletoDo For Append As fNum
    Print #fNum, "[ReaderStyle]"
    With RS.formPos
        Print #fNum, "FormHeight=" & CStr(.Height)
        Print #fNum, "FormWidth=" & CStr(.Width)
        Print #fNum, "FormTop=" & CStr(.Top)
        Print #fNum, "FormLeft=" & CStr(.Left)
    End With
    Print #fNum, "LastPath=" & RS.LastPath
    Print #fNum, "WindowState=" & CStr(RS.WindowState)
    Print #fNum, "LeftWidth=" & CStr(RS.LeftWidth)
    Print #fNum, "ShowMenu=" & CStr(RS.ShowMenu)
    Print #fNum, "ShowLeft=" & CStr(RS.ShowLeft)
    Print #fNum, "ShowStatusBar=" & CStr(RS.ShowStatusBar)
    Print #fNum, "FullScreenMode=" & CStr(RS.FullScreenMode)
    Print #fNum, "TextEditor=" & RS.TextEditor
    
    Close #fNum
End Sub

Public Sub GetViewerStyle(iniFiletoDo As String, VS As ViewerStyle)

    With VS.Viewfont
        .Bold = (Val(iniGetSetting(iniFiletoDo, "ViewStyle", "Bold")) > 0)
        .Italic = (Val(iniGetSetting(iniFiletoDo, "ViewStyle", "Italic")) > 0)
        .Underline = (Val(iniGetSetting(iniFiletoDo, "ViewStyle", "Underline")) > 0)
        .Strikethrough = (Val(iniGetSetting(iniFiletoDo, "ViewStyle", "Strikethrough")) > 0)
        .name = iniGetSetting(iniFiletoDo, "ViewStyle", "Name")
        .Size = Val(iniGetSetting(iniFiletoDo, "ViewStyle", "Size"))

        If .Size = 0 Then .Size = 9
    End With

    With VS
        .ForeColor = Val(iniGetSetting(iniFiletoDo, "ViewStyle", "ForeColor"))
        .BackColor = Val(iniGetSetting(iniFiletoDo, "ViewStyle", "BackColor"))
        .LineHeight = Val(iniGetSetting(iniFiletoDo, "ViewStyle", "LineHeight"))

        If .LineHeight = 0 Then .LineHeight = 100
    End With
    
    VS.RecentMax = Val(iniGetSetting(iniFiletoDo, "Viewstyle", "RecentMax"))

End Sub

Public Sub SaveViewerStyle(iniFiletoDo As String, VS As ViewerStyle)
    
    iniDeleteSection iniFiletoDo, "ViewStyle"
    
    Dim fNum As Integer
    fNum = FreeFile
    Open iniFiletoDo For Append As fNum
    Print #fNum, "[ViewStyle]"
    
    Dim a As Integer

    With VS.Viewfont

        If .Bold Then a = 1 Else a = 0
        Print #fNum, "Bold=" & CStr(a)

        If .Italic Then a = 1 Else a = 0
        Print #fNum, "Italic=" & CStr(a)

        If .Underline Then a = 1 Else a = 0
        Print #fNum, "Underline=" & CStr(a)

        If .Strikethrough Then a = 1 Else a = 0
        Print #fNum, "Strikethrough=" & CStr(a)
        Print #fNum, "Name=" & .name
        Print #fNum, "Size=" & CStr(.Size)
    End With

    With VS
        Print #fNum, "ForeColor=" & CStr(.ForeColor)
        Print #fNum, "Backcolor=" & CStr(.BackColor)
        Print #fNum, "LineHeight=" & CStr(.LineHeight)
        Print #fNum, "RecentMax=" & CStr(.RecentMax)
    End With
    
    Close #fNum
End Sub

Public Sub rememberNew(ByRef zhMemoryIn As String, ByVal szhFilename As String, ByVal ssecondPart As String)

    Dim fso As New Scripting.FileSystemObject
    Dim fsoMemoryTS As Scripting.TextStream
    Dim sMemoryText As String
    Dim stmp As String
    Dim posStart As Long
    Dim posEnd As Long
    Dim fMemoryDecrypted As String

    If szhFilename = "" Then Exit Sub
    If ssecondPart = "" Then Exit Sub

    fMemoryDecrypted = fso.BuildPath(Environ$("temp"), fso.GetTempName)
    MyFileDecrypt zhMemoryIn, fMemoryDecrypted
    Set fsoMemoryTS = fso.OpenTextFile(fMemoryDecrypted, ForReading, True)

    If fsoMemoryTS.AtEndOfStream = False Then sMemoryText = fsoMemoryTS.ReadAll
    fsoMemoryTS.Close
    Set fsoMemoryTS = fso.OpenTextFile(fMemoryDecrypted, ForWriting, True)
    posStart = InStr(sMemoryText, szhFilename & "|")

    If posStart > 0 Then posEnd = InStr(posStart, sMemoryText, zhMemorySplit, vbTextCompare)

    If posStart > 0 And posEnd > posStart Then
        stmp = Left$(sMemoryText, posStart - 1)
        stmp = stmp & szhFilename & "|" & ssecondPart & zhMemorySplit
        stmp = stmp & Right$(sMemoryText, Len(sMemoryText) - posEnd - Len(zhMemorySplit) + 1)
        sMemoryText = stmp
    Else
        sMemoryText = sMemoryText & szhFilename & "|" & ssecondPart & zhMemorySplit
    End If

    If Left$(sMemoryText, Len(zhMemorySplit)) = zhMemorySplit Then sMemoryText = Right$(sMemoryText, Len(sMemoryText) - Len(zhMemorySplit))
    fsoMemoryTS.Write sMemoryText
    fsoMemoryTS.Close
    MyFileEncrypt fMemoryDecrypted, zhMemoryIn
    fso.DeleteFile fMemoryDecrypted

End Sub

Public Function searchMemory(ByRef zhMemoryIn As String, ByRef szhFilename As String) As String

    Dim fso As New Scripting.FileSystemObject
    Dim fsoMemoryTS As Scripting.TextStream
    Dim sMemoryText As String
    Dim fMemoryDecrypted As String
    Dim posStart As Long
    Dim posEnd As Long

    fMemoryDecrypted = fso.BuildPath(Environ$("temp"), fso.GetTempName)
    MyFileDecrypt zhMemoryIn, fMemoryDecrypted
    Set fsoMemoryTS = fso.OpenTextFile(fMemoryDecrypted, ForReading, True)

    If fsoMemoryTS.AtEndOfStream = False Then sMemoryText = fsoMemoryTS.ReadAll
    fsoMemoryTS.Close
    posStart = InStr(sMemoryText, szhFilename & "|")

    If posStart > 0 Then posEnd = InStr(posStart, sMemoryText, zhMemorySplit, vbTextCompare)

    If posStart > 0 And posEnd > posStart + 1 Then
        searchMemory = Mid$(sMemoryText, posStart, posEnd - posStart)
        searchMemory = Replace(searchMemory, szhFilename & "|", "")
    End If

    fso.DeleteFile fMemoryDecrypted

End Function


Public Sub GetFileFilter(ByRef iniFiletoDo As String, ByRef cmbFilter As ComboBox)

Dim i As Integer
Dim ffNum As Long

ffNum = CLngStr(iniGetSetting(iniFiletoDo, "FileFilter", "Count")) - 1

For i = 0 To ffNum

cmbFilter.AddItem iniGetSetting(iniFiletoDo, "FileFilter", "F" + Str$(i)), i

Next

End Sub
Public Sub SaveFileFilter(ByRef iniFiletoDo As String, ByRef cmbFilter As ComboBox)

Dim ffNum As Long
Dim fNum As Integer
Dim i As Integer

iniDeleteSection iniFiletoDo, "FileFilter"

fNum = FreeFile
Open iniFiletoDo For Append As fNum
Print #fNum, "[FileFilter]"
Print #fNum, "Count=" & Str$(cmbFilter.ListCount)
ffNum = cmbFilter.ListCount - 1
For i = 0 To ffNum
Print #fNum, "F" & Str$(i) & "=" & cmbFilter.List(i)
Next
Close #fNum

End Sub

