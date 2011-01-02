Attribute VB_Name = "MainMOD"

Option Explicit

Type MYPoS
    Top As Long
    Left As Long
    Height As Long
    Width As Long
End Type

Public Enum zhtmVisiablity
zhtmVisiableTrue = 1
zhtmVisiableFalse = -1
zhtmVisiableDefault = 0
End Enum


Public Type zhReaderStatus
 sCur_zhFile As String
 sCur_zhSubFile As String
' bMenuShowed As Boolean
' bLeftShowed As Boolean
' bStatusBarShowed As Boolean
 sPWD As String
End Type

Public Type typeZhBookmark
sName As String
sZhfile As String
sZhsubfile As String
End Type

Public Type typeZhBookmarkCollection
Count As Integer
zhBookmark() As typeZhBookmark
End Type

Public Type ReadingStatus
page As String
perOfScrollTop As Single
perOfScrollLeft As Single
End Type


Public zhrStatus As zhReaderStatus
Public zhtmIni As String
Public LanguageIni As String
Public sTempZH As String
Public Tempdir As String
Public sConfigDir As String
Public Const TakeCare_EXT = "zpic"

Public Sub rememberBook(ByRef memFile As String, ByRef bookFile As String, ByRef nowAt As ReadingStatus)

    Dim hIni As New linvblib.CLiNInI
    Dim sectionName As String
    Dim fso As New FileSystemObject
    'If fso.FileExists(memFile) = False Then Exit Sub
    If fso.FileExists(bookFile) = False Then Exit Sub
    sectionName = fso.GetBaseName(bookFile) & "(" & CStr(FileLen(bookFile)) & ")"
    Set fso = Nothing
    
    On Error Resume Next
    hIni.Create memFile
    hIni.SaveSetting sectionName, "page", nowAt.page
    hIni.SaveSetting sectionName, "scrollTop", CStr(nowAt.perOfScrollTop)
    hIni.SaveSetting sectionName, "scrollLeft", CStr(nowAt.perOfScrollLeft)
    hIni.Save
       
    Set hIni = Nothing

End Sub
Public Function searchMem(ByRef memFile As String, ByRef bookFile As String) As ReadingStatus

    Dim hIni As New linvblib.CLiNInI
    Dim sectionName As String
    Dim fso As New FileSystemObject
    If fso.FileExists(memFile) = False Then Exit Function
    If fso.FileExists(bookFile) = False Then Exit Function
    sectionName = fso.GetBaseName(bookFile) & "(" & CStr(FileLen(bookFile)) & ")"
    Set fso = Nothing
    
    On Error Resume Next
    hIni.Create memFile
    With searchMem
    .page = hIni.GetSetting(sectionName, "page")
    .perOfScrollTop = CSng(hIni.GetSetting(sectionName, "scrollTop"))
    .perOfScrollLeft = CSng(hIni.GetSetting(sectionName, "scrollLeft"))
    End With
    Set hIni = Nothing

End Function



Private Sub Main()
             
             
    Association TakeCare_EXT
    
    Load MainFrm
    MainFrm.Show
    'startUP
    Dim thisfile As String
    thisfile = Command$

    If Left$(thisfile, "1") = Chr$(34) And Right$(thisfile, 1) = Chr$(34) And Len(thisfile) > 1 Then
        thisfile = Right$(thisfile, Len(thisfile) - 1)
        thisfile = Left$(thisfile, Len(thisfile) - 1)
    End If

    If thisfile <> "" Then
        MainFrm.loadzh thisfile
    Else
        MainFrm.appAbout
    End If

End Sub

Public Function htmlline(ByRef text) As String
htmlline = Replace$(text, "!", Chr$(34))
End Function

Private Function Association(ByRef strExtName) As Boolean
    Dim hReg As New CRegistry
    
    hReg.ClassKey = HKEY_CLASSES_ROOT
    hReg.SectionKey = "zpicfile"
    
    If hReg.KeyExists = True Then
        hReg.SectionKey = "." & strExtName
        hReg.Value = "zhtmfile"
        Association = False
        Exit Function
    End If
    
    
    hReg.CreateEXEAssociation _
        bddir(App.Path) & App.EXEName & ".exe", _
        "zpicfile", _
        "Zip archive of pictures", _
        strExtName, _
        , False, , False, , False, "", 3
    Set hReg = Nothing
    
    Association = True
    
End Function

Private Function AddAssociation(ByRef strExtName) As Boolean
End Function
