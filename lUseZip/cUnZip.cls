VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUnzip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'©°©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤
'©¦     Component  : cUnzip
'©¦     Project    : LUseZipDll
'©¦     Description: Use to Unzip zipFile Through Info-zip Unzip V.5.40
'©¦
'©¦
'©¦     Modified   : xrLin
'©¦     Date       : 2005-11-9
'©¸©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤

'©°©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤
'©¦     Name:     vbAccelerator cUnzip class
'©¦     Author:   Steve McMahon (steve@vbaccelerator.com)
'©¦     Date:     1 December 2000
'©¦
'©¦     Requires: Info-ZIP'©¦s Unzip32.DLL v5.40, renamed to vbuzip10.dll
'©¦               mUnzip.bas
'©¦
'©¦     Copyright © 2000 Steve McMahon for vbAccelerator
'©¦
'©¦     Visit vbAccelerator - advanced free source code for VB programmers
'©¦     http://vbaccelerator.com
'©¸©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤

'©°©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤
'©¦     Part of the implementation of cUnzip.cls, a class which gives a
'©¦     simple interface to Info-ZIP'©¦s excellent, free unzipping library
'©¦     (Unzip32.DLL).
'©¦
'©¸©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤

Option Explicit

Public Enum EUZMsgLevel
   euzAllMessages = 0
   euzPartialMessages = 1
   euzNoMessages = 2
End Enum
Public Enum EUZOverWriteResponse
   euzDoNotOverwrite = 100
   euzOverwriteThisFile = 102
   euzOverwriteAllFiles = 103
   euzOverwriteNone = 104
End Enum
Public Enum unzReturnCode
 PK_OK = 0               '/* no error */
 PK_COOL = 0             '/* no error */
 PK_WARN = 1             '/* warning error */
 PK_ERR = 2              '/* error in zipfile */
 PK_BADERR = 3           '/* severe error in zipfile */
 PK_MEM = 4              '/* insufficient memory (during initialization) */
 PK_MEM2 = 5             '/* insufficient memory (password failure) */
 PK_MEM3 = 6             '/* insufficient memory (file decompression) */
 PK_MEM4 = 7             '/* insufficient memory (memory decompression) */
 PK_MEM5 = 8             '/* insufficient memory (not yet used) */
 PK_NOZIP = 9            '/* zipfile not found */
 PK_PARAM = 10           '/* bad or illegal parameters specified */
 PK_FIND = 11            '/* no files found */
 PK_DISK = 50            '/* disk full */
 PK_EOF = 51             '/* unexpected EOF */
 IZ_CTRLC = 80           '/* user hit ^C to terminate */
 IZ_UNSUP = 81           '/* no files found: all unsup. compr/encrypt. */
 IZ_BADPWD = 82          '/* no files found: all had bad password */
End Enum
Private m_sZipFile As String
Private m_sUnzipFolder As String
Private m_ZipItems As CZipItems
Private m_iCount As Long
Private m_tDCL As DCLIST
Private m_sFileToProcess() As String
Private m_sFTPcount As Long
Private m_sFileToExclude() As String
Private m_sFTEcount As Long
Private m_bGetComment As Boolean
Private m_sComment As String
Private m_sFilelistSeparator As String
Private Const mc_sFileListSeparator = "|"
Public Event Cancel(ByVal sMsg As String, ByRef bCancel As Boolean)
Public Event OverwritePrompt(ByVal sFile As String, ByRef eResponse As EUZOverWriteResponse)
Public Event PasswordRequest(ByRef sPassword As String, ByVal sName As String, ByRef bCancel As Boolean)
Public Event Progress(ByVal lCount As Long, ByVal sMsg As String)
'local variable(s) to hold property value(s)
Private mvarisGetZipItems As Boolean 'local copy
Friend Property Let isGetZipItems(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.isGetZipItems = 5
    mvarisGetZipItems = vData
End Property


Friend Property Get isGetZipItems() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.isGetZipItems
    isGetZipItems = mvarisGetZipItems
End Property




Friend Property Get isGetComment() As Boolean
isGetComment = m_bGetComment
End Property
Friend Property Let isGetComment(TrueOrFalse As Boolean)
m_bGetComment = TrueOrFalse
End Property

Public Property Get ExtractOnlyNewer() As Boolean
   ExtractOnlyNewer = (m_tDCL.ExtractOnlyNewer <> 0)      ' 1=extract only newer
End Property
Public Property Let ExtractOnlyNewer(ByVal bState As Boolean)
   m_tDCL.ExtractOnlyNewer = Abs(bState)      ' 1=extract only newer
End Property
Public Property Get SpaceToUnderscore() As Boolean
   SpaceToUnderscore = (m_tDCL.SpaceToUnderscore <> 0)    ' 1=convert space to underscore
End Property
Public Property Let SpaceToUnderscore(ByVal bState As Boolean)
   m_tDCL.SpaceToUnderscore = Abs(bState)
End Property
Public Property Get PromptToOverwrite() As Boolean
   PromptToOverwrite = (m_tDCL.PromptToOverwrite <> 0)
End Property
Public Property Let PromptToOverwrite(ByVal bState As Boolean)
   m_tDCL.PromptToOverwrite = Abs(bState)
End Property
Public Property Get MessageLevel() As EUZMsgLevel
   MessageLevel = m_tDCL.fQuiet
End Property
Public Property Let MessageLevel(ByVal eLevel As EUZMsgLevel)
   m_tDCL.fQuiet = eLevel
End Property
Public Property Get TestZip() As Boolean
   TestZip = (m_tDCL.ntflag <> 0)
End Property
Public Property Let TestZip(ByVal bState As Boolean)
   m_tDCL.ntflag = Abs(bState)
End Property
Public Property Get UseFolderNames() As Boolean
   UseFolderNames = (m_tDCL.ndflag <> 0)
End Property
Public Property Let UseFolderNames(ByVal bState As Boolean)
   m_tDCL.ndflag = Abs(bState)
End Property
Public Property Get OverwriteExisting() As Boolean
   OverwriteExisting = (m_tDCL.noflag <> 0)
End Property
Public Property Let OverwriteExisting(ByVal bState As Boolean)
   m_tDCL.noflag = Abs(bState)
End Property
Public Property Get ConvertCRToCRLF() As Boolean
   ConvertCRToCRLF = (m_tDCL.naflag <> 0)
End Property
Public Property Let ConvertCRToCRLF(ByVal bState As Boolean)
   m_tDCL.naflag = Abs(bState)
End Property
Public Property Get CaseSensitiveFileNames() As Boolean
   CaseSensitiveFileNames = m_tDCL.C_flag
End Property
Public Property Let CaseSensitiveFileNames(ByVal bState As Boolean)
   m_tDCL.C_flag = Abs(bState)
End Property

Friend Sub AddZipItem( _
      ByVal sFilename As String, _
      ByVal dDate As Date, _
      ByVal lUSize As Long, _
      ByVal lSize As Long, _
      ByVal lCrc As Long, _
      ByVal fEncrypted As Boolean, _
      ByVal lFactor As Long, _
      ByVal sMethod As String _
      )
'   MsgBox "EnterMessage " & sFilename
   On Error Resume Next
   
   If (sFilename <> vbNullChar) And Len(sFilename) > 0 Then
   
   Debug.Print sFilename, , , , lSize; "B", , , , lFactor; "%"
   m_ZipItems.Add sFilename, lUSize, lSize, lFactor, sMethod, dDate, lCrc, fEncrypted
   

'      m_iCount = m_iCount + 1
'      ReDim Preserve m_tZipContents(1 To m_iCount) As tZipContents
'      With m_tZipContents(m_iCount)
'         .sFilename = sFilename
'         .dDate = dDate
'         .lSize = lSize
'         .lCrc = lCrc
'         .lFactor = lFactor
'         .sMethod = sMethod
'         .fEncryped = fEncrypted
'         ' Default to selected:
'         '.fSelected = True
'
'      End With
   End If

End Sub
Friend Sub OverwriteRequest( _
      ByVal sFile As String, _
      ByRef eResponse As EUZOverWriteResponse _
   )
   RaiseEvent OverwritePrompt(sFile, eResponse)
End Sub
Friend Sub ProgressReport( _
      ByVal sMsg As String _
   )
'   Debug.Print sMsg
   RaiseEvent Progress(1, sMsg)
End Sub
Friend Sub Service( _
      ByVal sMsg As String, _
      ByRef bCancel As Boolean _
   )
   RaiseEvent Cancel(sMsg, bCancel)
End Sub
      
Friend Sub PasswordRequest( _
      ByRef sPassword As String, _
      ByRef sName As String, _
      ByRef bCancel As Boolean _
   )
   RaiseEvent PasswordRequest(sPassword, sName, bCancel)
End Sub
'

Public Function getZipItems(ByRef ItemsPlaceHere As CZipItems) As unzReturnCode
    
    Dim S(0 To 0) As String
    Dim tempDCL As DCLIST
    Dim tmpZipItem As CZipItem
    Dim ret As unzReturnCode
    
    With tempDCL
       .lpszZipFN = m_sZipFile
       .lpszExtractDir = vbNullChar
       .nvflag = 1
    End With
    

    Set m_ZipItems = New CZipItems
    
    isGetZipItems = True
    ret = mUnzip.VBUnzip(Me, tempDCL, 0, S(), 0, S())
    isGetZipItems = False
        
    For Each tmpZipItem In m_ZipItems
        With tmpZipItem
            ItemsPlaceHere.Add .Filename, .Size, .PackedSize, .Factor, .Method, .CreatedDate, .Crc, .Encrypted
        End With
    Next
       
    Set tmpZipItem = Nothing
    Set m_ZipItems = Nothing
        
End Function

'Public Function Directory() As unzReturnCode
'Dim S(0 To 0) As String
'Dim tempDCL As DCLIST
'With tempDCL
'   .lpszZipFN = m_sZipFile
'   .lpszExtractDir = vbNullChar
'   .nvflag = 1
'End With
'  Directory = mUnzip.VBUnzip(Me, tempDCL, 0, S(), 0, S())
'End Function
Public Function unzip() As unzReturnCode
   unzip = PK_FIND
   If (m_sZipFile <> "") Then
      m_tDCL.C_flag = 1
      m_tDCL.lpszZipFN = m_sZipFile
      m_tDCL.nvflag = 0
      m_tDCL.lpszExtractDir = m_sUnzipFolder
      unzip = mUnzip.VBUnzip(Me, m_tDCL, m_sFTPcount, m_sFileToProcess, m_sFTEcount, m_sFileToExclude())
   End If
End Function
'Private Function ReverseSlashes( _
'      ByVal sFolder As String, _
'      ByVal sFile As String _
'   ) As String
'Dim sOut As String
'Dim iPos As Long, iLastPos As Long
'
'   If Len(sFolder) > 0 And sFolder <> vbNullChar Then
'      sOut = sFolder & "/" & sFile
'      iLastPos = 1
'      Do
'         iPos = InStr(iLastPos, sOut, "\")
'         If (iPos <> 0) Then
'            Mid$(sOut, iPos, 1) = "/"
'            iLastPos = iPos + 1
'         End If
'      Loop While iPos <> 0
'      ReverseSlashes = sOut
'   Else
'      ReverseSlashes = sFile
'   End If
'End Function
Public Property Let UnzipFolder(ByVal sFolder As String)
   m_sUnzipFolder = sFolder
End Property
Public Property Get UnzipFolder() As String
   UnzipFolder = m_sUnzipFolder
   m_tDCL.lpszExtractDir = m_sUnzipFolder
End Property
Public Property Get ZipFile() As String
   ZipFile = m_sZipFile
End Property
'Public Sub GetVersion( _
'      ByRef lMajor As Long, _
'      ByRef lMinor As Long, _
'      ByRef lRevision As Long, _
'      ByRef dDate As Date, _
'      ByRef sExtraInfo As String _
'   )
'Dim tVer As UZPVER
'Dim iPos As Long
'
'   ' Set Version space
'   With tVer
'      .structlen = Len(tVer)
'      .betalevel = Space$(9) & vbNullChar
'      .date = Space$(19) & vbNullChar
'      .zlib = Space$(9) & vbNullChar
'   End With
'
'   ' Get version
'   UzpVersion2 tVer
'   iPos = InStr(tVer.betalevel, vbNullChar)
'   If (iPos > 1) Then
'      tVer.betalevel = Left$(tVer.betalevel, iPos - 1)
'   End If
'   sExtraInfo = tVer.betalevel
'
'   ' Date..
'   Debug.Print tVer.date
'
'   lMajor = tVer.windll.major
'   lMinor = tVer.windll.minor
'   lRevision = tVer.windll.patchlevel
'
'End Sub

Public Property Let ZipFile(ByVal sFile As String)
   m_sZipFile = sFile
'   m_iCount = 0
'   Erase m_tZipContents
End Property

Public Property Let FileToProcess(ByVal sFileList As String)
m_sFileToProcess = Split(sFileList, FileListSeparator)
m_sFTPcount = UBound(m_sFileToProcess) + 1
End Property
Public Property Get FileToProcess() As String
Dim l As Long
For l = 0 To m_sFTPcount - 1
FileToProcess = FileToProcess & m_sFileToProcess(l) & vbCrLf
Next
End Property
Public Property Get FileListSeparator() As String
    If m_sFilelistSeparator = "" Then m_sFilelistSeparator = mc_sFileListSeparator
    FileListSeparator = m_sFilelistSeparator
End Property
Public Property Let FileListSeparator(ByVal sValue As String)
    FileListSeparator = sValue
End Property

Public Property Let FileToExclude(ByVal sFileList As String)
m_sFileToExclude = Split(sFileList, FileListSeparator)
m_sFTEcount = UBound(m_sFileToExclude) + 1
End Property
Public Property Get FileToExclude() As String
Dim l As Long
For l = 0 To m_sFTEcount - 1
FileToExclude = FileToExclude & m_sFileToExclude(l) & vbCrLf
Next
End Property
Public Sub AddFileToPreocess(ByVal sFilename As String)
If sFilename = "" Then Exit Sub
ReDim m_sFileToProcess(m_sFTPcount)
m_sFileToProcess(m_sFTPcount) = sFilename
m_sFTPcount = m_sFTPcount + 1
End Sub
Public Sub AddFileToExclude(ByVal sFilename As String)
If sFilename = "" Then Exit Sub
ReDim m_sFileToExclude(m_sFTEcount)
m_sFileToExclude(m_sFTEcount) = sFilename
m_sFTEcount = m_sFTEcount + 1
End Sub
Public Function GetComment() As String
If m_sZipFile = "" Then Exit Function

Dim tempDCL As DCLIST
Dim S(0 To 0) As String

With tempDCL
   .nzflag = 1
   .ntflag = 1
   .lpszZipFN = m_sZipFile
   .lpszExtractDir = vbNullChar
End With
   'm_tDCL.nvflag = 1
   Me.Comment = "'"
   isGetComment = True
   mUnzip.VBUnzip Me, tempDCL, 0, S(), 0, S()
   isGetComment = False
   GetComment = Me.Comment
   
End Function
Public Property Let Comment(sComment As String)
m_sComment = sComment
End Property

Public Property Get Comment() As String
Comment = m_sComment
End Property

Public Function ValidateZipFile(sZipfilename As String) As Boolean
If Wiz_Validate(ByVal sZipfilename, 1) = PK_OK Then ValidateZipFile = True
End Function


