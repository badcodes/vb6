VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IUnzip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
''========================================
''Unzip
''========================================
Public Enum MPUnzipMsgLevel
   MP_Unzip_AllMessages ' = 0
   MP_Unzip_PartialMessages ' = 1
   MP_Unzip_NoMessages ' = 2
End Enum
Public Enum MPUnzipReturnCode
 MP_Unzip_OK ' = 0               '/* no error */
 MP_Unzip_COOL ' = 0             '/* no error */
 MP_Unzip_WARN ' = 1             '/* warning error */
 MP_Unzip_ERR ' = 2              '/* error in zipfile */
 MP_Unzip_BADERR ' = 3           '/* severe error in zipfile */
 MP_Unzip_MEM ' = 4              '/* insufficient memory (during initialization) */
 MP_Unzip_MEM2 ' = 5             '/* insufficient memory (password failure) */
 MP_Unzip_MEM3 ' = 6             '/* insufficient memory (file decompression) */
 MP_Unzip_MEM4 ' = 7             '/* insufficient memory (memory decompression) */
 MP_Unzip_MEM5 ' = 8             '/* insufficient memory (not yet used) */
 MP_Unzip_NOZIP ' = 9            '/* zipfile not found */
 MP_Unzip_PARAM ' = 10           '/* bad or illegal parameters specified */
 MP_Unzip_FIND ' = 11            '/* no files found */
 MP_Unzip_DISK ' = 50            '/* disk full */
 MP_Unzip_EOF ' = 51             '/* unexpected EOF */
 MP_Unzip_CTRLC ' = 80           '/* user hit ^C to terminate */
 MP_Unzip_UNSUP ' = 81           '/* no files found: all unsup. compr/encrypt. */
 MP_Unzip_BADPWD ' = 82          '/* no files found: all had bad password */
End Enum
'Public Event Cancel(ByVal sMsg As String, ByRef bCancel As Boolean)
'Public Event OverwritePrompt(ByVal sFile As String, ByRef eResponse As EUZOverWriteResponse)
'Public Event PasswordRequest(ByRef sPassword As String, ByVal sName As String, ByRef bCancel As Boolean)
'Public Event Progress(ByVal lCount As Long, ByVal sMsg As String)

Public Sub AddFilesToExclude(ByVal strFilename As String, Optional strListSeparator As String = ",")
End Sub

Public Sub AddFileToPreocess(ByVal sFilename As String, Optional strListSeparator As String = ",")
End Sub

Public Property Let CaseSensitive(ByVal bValue As Boolean)
End Property

Public Property Get CaseSensitive() As Boolean
End Property

Public Property Let Comment(ByVal sValue As String)
End Property

Public Property Get Comment() As String
End Property

Public Property Let ConvertCRToCRLF(ByVal bValue As Boolean)
End Property

Public Property Get ConvertCRToCRLF() As Boolean
End Property

Public Property Let ExtractOnlyNewer(ByVal bValue As Boolean)
End Property

Public Property Get ExtractOnlyNewer() As Boolean
End Property

Public Property Let FilesToExclude(ByVal sValue As String)
End Property

Public Property Get FilesToExclude() As String
End Property

Public Property Let FilesToProcess(ByVal sValue As String)
End Property

Public Property Get FilesToProcess() As String
End Property

Public Function getZipItems(ByRef zipItems As IArchiveItems) As MPUnzipReturnCode
End Function

Public Property Get MessageLevel() As MPUnzipMsgLevel
End Property

Public Property Let MessageLevel(ByVal msgLvl As MPUnzipMsgLevel)
End Property

Public Property Let OverwriteExisting(ByVal bValue As Boolean)
End Property

Public Property Get OverwriteExisting() As Boolean
End Property

Public Property Let PromptToOverwrite(ByVal bValue As Boolean)
End Property

Public Property Get PromptToOverwrite() As Boolean
End Property
Public Property Let UnzipToFolder(ByVal sValue As String)
End Property
Public Property Get UnzipToFolder() As String
End Property
Public Property Let UseFolderNames(ByVal bValue As Boolean)
End Property
Public Property Get UseFolderNames() As Boolean
End Property
Public Function Test() As Boolean
End Function
Public Function Validate() As Boolean
End Function
Public Function GetComment() As String
End Function
Public Function Unzip() As MPUnzipReturnCode
End Function
Public Property Let Filename(ByVal sValue As String)
End Property
Public Property Get Filename() As String
End Property



