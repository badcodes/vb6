VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gMSSReader"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum SSLIBTaskStatus
    STS_START = 1
    STS_PAUSE = 0
    STS_COMPLETE = 2
    STS_PENDING = 3
    STS_ERRORS = 4
End Enum

Private Const CST_SSLIB_FIELDS_LBound As Long = 1

Public Enum SSLIBFields
        SSF_Title = CST_SSLIB_FIELDS_LBound
        SSF_AUTHOR
        SSF_SSID
        SSF_ISJPGBOOK
        SSF_PagesCount
        SSF_Publisher
        SSF_PublishDate
        SSF_StartPage
        
        SSF_Subject
        SSF_Comments
        
        SSF_URL
        SSF_JPGURL
        SSF_SAVEDIN
        SSF_HEADER
        'SSF_FULLNAME
        
        SSF_Downloader
        SSF_DownloadDate
        'SSF_STATUS
        'SSF_FILES_DOWNLOADED
        
        SSF_PageURL
        SSF_HTMLContent
        SSF_FIELDS_END

End Enum
Private Const CST_SSLIB_FIELDS_UBound As Long = SSF_FIELDS_END - 1
Private Const CST_SSLIB_FIELDS_IMPORTANT_UBOUND As Long = SSF_Comments
Private Const CST_SSLIB_FIELDS_TASKS_UBOUND As Long = SSF_HEADER

