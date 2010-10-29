VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum EZPMsgLevel
   ezpAllMessages = 0
   ezpPartialMessages = 1
   ezpNoMessages = 2
End Enum

Public Enum zpCompressionLevel
    zpLevel0_Store = 0
    zpLevel1 = 1
    zpLevel2 = 2
    zpLevel3 = 3
    zpLevel4 = 4
    zpLevel5 = 5
    zpLevel6_Default = 6
    zpLevel7 = 7
    zpLevel8 = 8
    zpLevel9_MAX = 9
End Enum

Public Event Cancel(ByVal sMsg As String, ByRef bCancel As Boolean)
Public Event PasswordRequest(ByRef sPassword As String, ByRef bCancel As Boolean)
Public Event Progress(ByVal lCount As Long, ByVal sMsg As String)

Private m_tZPOPT As ZPOPT
Private m_sFilename As String
'Private m_sFileSpecs() As String
Private m_sComment As String
'Private m_iCount As Long
Private m_sFileToProcess() As String
Private m_sFTPcount As Long
Private m_sFilelistSeparator As String
Private Const mc_sFileListSeparator = "|"

' Set zip options
' m_tZPOPT.fSuffix = 0        ' include suffixes (not yet implemented)
' m_tZPOPT.fExtra = 0         ' 1 if including extra attributes
' m_tZPOPT.date = vbNullString ' "12/31/79"? US Date?
' m_tZPOPT.fExcludeDate = 0   ' 1 if excluding files earlier than a specified date
' m_tZPOPT.fIncludeDate = 0   ' 1 if including files earlier than a specified date
' m_tZPOPT.fJunkSFX = 0       ' 1 if junking sfx prefix
' m_tZPOPT.fOffsets = 0       ' 1 if updating archive offsets for sfx Files

' m_tZPOPT.fComment = 0       ' 1 if putting comment in zip file

' m_tZPOPT.fGrow = 0          ' 1 if allow appending to zip file
' m_tZPOPT.fForce = 0         ' 1 if making entries using DOS names
' m_tZPOPT.fMove = 0          ' 1 if deleting files added or updated
' m_tZPOPT.fDeleteEntries = 0 ' 1 if files passed have to be deleted
' m_tZPOPT.fLatestTime = 0    ' 1 if setting zip file time to time of latest file in archive
' m_tZPOPT.fPrivilege = 0     ' 1 if not saving privileges
' m_tZPOPT.fEncryption = 0    'Read only property!
' m_tZPOPT.fRepair = 0        ' 1=> fix archive, 2=> try harder to fix
' m_tZPOPT.flevel = 0         ' compression level - should be 0!!!
Public Property Get SaveComment() As Boolean
SaveComment = Not (m_tZPOPT.fComment = 0)
End Property
Public Property Let SaveComment(bState As Boolean)
m_tZPOPT.fComment = Abs(bState)
End Property

Public Property Get CompressLevel() As zpCompressionLevel
    CompressLevel = CInt(Chr$(m_tZPOPT.fLevel))
End Property
Public Property Let CompressLevel(CompressLevelValue As zpCompressionLevel)
    m_tZPOPT.fLevel = Asc(LTrim$(CStr(CompressLevelValue)))
End Property

Public Property Get ZipFile() As String
   ZipFile = m_sFilename
End Property
Public Property Let ZipFile(ByVal sFilename As String)
   m_sFilename = sFilename
End Property
Public Property Get BasePath() As String
   BasePath = m_tZPOPT.szRootDir
End Property
Public Property Let BasePath(ByVal sBasePath As String)
   m_tZPOPT.szRootDir = sBasePath
End Property
Public Property Get Encrypt() As Boolean
   Encrypt = Not (m_tZPOPT.fEncrypt = 0)
End Property
Public Property Let Encrypt(ByVal bState As Boolean)
   m_tZPOPT.fEncrypt = Abs(bState)
End Property
Public Property Get IncludeSystemAndHiddenFiles() As Boolean
    IncludeSystemAndHiddenFiles = Not (m_tZPOPT.fSystem = 0)       ' 1 to include system/hidden files
End Property
Public Property Let IncludeSystemAndHiddenFiles(ByVal bState As Boolean)
   m_tZPOPT.fSystem = Abs(bState)       ' 1 to include system/hidden files
End Property
Public Property Get StoreVolumeLabel() As Boolean
    StoreVolumeLabel = Not (m_tZPOPT.fVolume = 0)       ' 1 if storing volume label
End Property
Public Property Let StoreVolumeLabel(ByVal bState As Boolean)
   m_tZPOPT.fVolume = Abs(bState)
End Property
Public Property Get StoreDirectories() As Boolean
   StoreDirectories = Not (m_tZPOPT.fNoDirEntries = 0) ' 1 if ignoring directory entries
End Property
Public Property Let StoreDirectories(ByVal bState As Boolean)
   m_tZPOPT.fNoDirEntries = Abs(Not (bState))
End Property
Public Property Get StoreFolderNames() As Boolean
   StoreFolderNames = (m_tZPOPT.fJunkDir = 0)
End Property
Public Property Let StoreFolderNames(ByVal bState As Boolean)
   m_tZPOPT.fJunkDir = Abs(Not (bState))
End Property
Public Property Get RecurseSubDirs() As Boolean
   RecurseSubDirs = Not (m_tZPOPT.fRecurse = 0) ' 1 if recursing into subdirectories
End Property
Public Property Let RecurseSubDirs(ByVal bState As Boolean)
   If bState Then
      m_tZPOPT.fRecurse = 2
   Else
      m_tZPOPT.fRecurse = 0
   End If
End Property

Public Property Get UpdateOnlyIfNewer() As Boolean
    UpdateOnlyIfNewer = Not (m_tZPOPT.fUpdate = 0)   ' 1 if updating zip file--overwrite only if newer
End Property
Public Property Let UpdateOnlyIfNewer(ByVal bState As Boolean)
    m_tZPOPT.fUpdate = Abs(bState)   ' 1 if updating zip file--overwrite only if newer
End Property
Public Property Get FreshenFiles() As Boolean
    FreshenFiles = Not (m_tZPOPT.fFreshen = 0) ' 1 if freshening zip file--overwrite only
End Property
Public Property Let FreshenFiles(ByVal bState As Boolean)
    m_tZPOPT.fUpdate = Abs(bState)   ' 1 if updating zip file--overwrite only if newer
End Property
Public Property Get MessageLevel() As EZPMsgLevel
   If Not (m_tZPOPT.fVerbose = 0) Then
      MessageLevel = ezpAllMessages
   ElseIf Not (m_tZPOPT.fQuiet = 0) Then
      MessageLevel = ezpPartialMessages
   Else
      MessageLevel = ezpNoMessages
   End If
End Property
Public Property Let MessageLevel(ByVal eLevel As EZPMsgLevel)
   Select Case eLevel
   Case ezpPartialMessages
      m_tZPOPT.fQuiet = 1
      m_tZPOPT.fVerbose = 0
   Case ezpNoMessages
      m_tZPOPT.fQuiet = 0
      m_tZPOPT.fVerbose = 0
   Case ezpAllMessages
      m_tZPOPT.fQuiet = 0
      m_tZPOPT.fVerbose = 1
   End Select
End Property
Public Property Get ConvertCRLFToLF() As Boolean
   ConvertCRLFToLF = (m_tZPOPT.fCRLF_LF <> 0)
End Property
Public Property Let ConvertCRLFToLF(ByVal bState As Boolean)
   m_tZPOPT.fCRLF_LF = Abs(bState)
End Property
Public Property Get ConvertLFToCRLF() As Boolean
   ConvertLFToCRLF = (m_tZPOPT.fLF_CRLF <> 0)
End Property
Public Property Let ConvertLFToCRLF(ByVal bState As Boolean)
   m_tZPOPT.fLF_CRLF = Abs(bState)
End Property

Friend Sub ProgressReport( _
      ByVal sMsg As String _
   )
   RaiseEvent Progress(1, sMsg)
End Sub
Friend Sub PasswordRequest( _
      ByRef sPassword As String, _
      ByRef bCancel As Boolean _
   )
   RaiseEvent PasswordRequest(sPassword, bCancel)
End Sub
Friend Sub Service( _
      ByVal sMsg As String, _
      ByRef bCancel As Boolean _
   )
   RaiseEvent Cancel(sMsg, bCancel)
End Sub
Public Property Get AllowAppend() As Boolean
   AllowAppend = (m_tZPOPT.fGrow = 1)
End Property
Public Property Let AllowAppend(ByVal bState As Boolean)
   m_tZPOPT.fGrow = Abs(bState)
End Property
Public Sub Zip()
   mZip.VBZip Me, m_tZPOPT, m_sFileToProcess(), m_sFTPcount
End Sub
Public Sub ZipComment(sComment As String)
           
   Dim sTmpFile(0) As String
   Dim tempZPOPT As ZPOPT
   If Me.ZipFile = "" Then Exit Sub
   m_sComment = sComment
   tempZPOPT.fComment = 1
   mZip.VBZip Me, tempZPOPT, sTmpFile(), 0
   
End Sub
Public Sub Delete()
   ' Deletes the entries specified by the file specs:
   m_tZPOPT.fDeleteEntries = 1
   'm_tZPOPT.fMove = 1
   mZip.VBZip Me, m_tZPOPT, m_sFileToProcess(), m_sFTPcount
  ' m_tZPOPT.fMove = 0
   m_tZPOPT.fDeleteEntries = 0
   
End Sub

Private Sub Class_Initialize()
   With m_tZPOPT
   .fPrivilege = 1
   End With
   m_tZPOPT.szRootDir = vbNullString
   StoreDirectories = True
   StoreFolderNames = False
   RecurseSubDirs = False
   CompressLevel = zpLevel6_Default
End Sub

Public Property Let Comment(sComment As String)
m_sComment = sComment
End Property

Public Property Get Comment() As String
Comment = m_sComment
End Property
Public Property Let FileToProcess(ByVal sFileList As String)
m_sFileToProcess = Split(sFileList, FileListSeparator)
m_sFTPcount = UBound(m_sFileToProcess) + 1
End Property
Public Property Get FileToProcess() As String
Dim l As Long
Dim lEnd As Long
lEnd = m_sFTPcount - 1
For l = 0 To lEnd
FileToProcess = FileToProcess & m_sFileToProcess(l) & vbCrLf
Next
End Property
Public Sub AddFileToProcess(ByVal sFilename As String)

If sFilename = "" Then Exit Sub
ReDim m_sFileToProcess(m_sFTPcount)
m_sFileToProcess(m_sFTPcount) = sFilename
m_sFTPcount = m_sFTPcount + 1

End Sub
Public Property Get FileListSeparator() As String
    If m_sFilelistSeparator = "" Then m_sFilelistSeparator = mc_sFileListSeparator
    FileListSeparator = m_sFilelistSeparator
End Property
Public Property Let FileListSeparator(ByVal sValue As String)
    FileListSeparator = sValue
End Property
