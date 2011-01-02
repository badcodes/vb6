VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements IUnzip

Private Sub IUnzip_AddFilesToExclude(ByVal strFilename As String, Optional strListSeparator As String = ",")
End Sub

Private Sub IUnzip_AddFileToPreocess(ByVal sFilename As String, Optional strListSeparator As String = ",")
End Sub

Private Property Get IUnzip_CaseSensitive() As Boolean
End Property

Private Property Let IUnzip_CaseSensitive(ByVal RHS As Boolean)

End Property

Private Property Let IUnzip_Comment(ByVal RHS As String)

End Property

Private Property Get IUnzip_Comment() As String

End Property

Private Property Let IUnzip_ConvertCRToCRLF(ByVal RHS As Boolean)

End Property

Private Property Get IUnzip_ConvertCRToCRLF() As Boolean

End Property

Private Property Let IUnzip_ExtractOnlyNewer(ByVal RHS As Boolean)

End Property

Private Property Get IUnzip_ExtractOnlyNewer() As Boolean

End Property

Private Property Let IUnzip_Filename(ByVal RHS As String)

End Property

Private Property Get IUnzip_Filename() As String

End Property

Private Property Let IUnzip_FilesToExclude(ByVal RHS As String)

End Property

Private Property Get IUnzip_FilesToExclude() As String

End Property

Private Property Let IUnzip_FilesToProcess(ByVal RHS As String)

End Property

Private Property Get IUnzip_FilesToProcess() As String

End Property

Private Function IUnzip_GetComment() As String

End Function

Private Function IUnzip_getZipItems(ItemsPlaceHere As LUseZipDll.CZipItems) As LUseZipDll.unzReturnCode

End Function

Private Property Let IUnzip_MessageLevel(ByVal RHS As LUseZipDll.EUZMsgLevel)

End Property

Private Property Get IUnzip_MessageLevel() As LUseZipDll.EUZMsgLevel

End Property

Private Property Let IUnzip_OverwriteExisting(ByVal RHS As Boolean)

End Property

Private Property Get IUnzip_OverwriteExisting() As Boolean

End Property

Private Property Let IUnzip_PromptToOverwrite(ByVal RHS As Boolean)

End Property

Private Property Get IUnzip_PromptToOverwrite() As Boolean

End Property

Private Function IUnzip_Test() As Boolean

End Function

Private Function IUnzip_Unzip() As LUseZipDll.unzReturnCode

End Function

Private Property Let IUnzip_UnzipToFolder(ByVal RHS As String)

End Property

Private Property Get IUnzip_UnzipToFolder() As String

End Property

Private Property Let IUnzip_UseFolderNames(ByVal RHS As Boolean)

End Property

Private Property Get IUnzip_UseFolderNames() As Boolean

End Property

Private Function IUnzip_Validate() As Boolean

End Function
Public Property Get InterfaceOfUnzip() As IUnzip
    Set InterfaceUnzip = Me
End Property
