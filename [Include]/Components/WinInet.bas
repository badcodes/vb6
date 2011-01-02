Attribute VB_Name = "mdWinInet"
Option Explicit
'--- Consts
Private Const MODULE_NAME           As String = "mdWinInet"
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const FTP_TRANSFER_TYPE_ASCII = &H1                          '--- 0x00000001
Public Const FTP_TRANSFER_TYPE_BINARY = &H2                         '--- 0x00000002
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000                      '--- 0x20000000
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000               '--- 0x10000000
Public Const HTTP_ADDREQ_FLAG_COALESCE = &H40000000                 '--- 0x40000000
Public Const HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA = &H40000000      '--- 0x40000000
Public Const HTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON = &H1000000   '--- 0x01000000
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000                  '--- 0x80000000
Public Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000              '--- 0x00010000
Public Const INTERNET_FLAG_DONT_CACHE = &H4000000                   '--- 0x04000000
Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000               '--- 0x04000000
Public Const INTERNET_FLAG_HYPERLINK = &H400                        '--- 0x00000400
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000          '--- 0x00001000
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000        '--- 0x00001000
Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000         '--- 0x00008000
Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000        '--- 0x00004000
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Public Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000
Public Const INTERNET_FLAG_MUST_CACHE_REQUEST = &H10
Public Const INTERNET_FLAG_NEED_FILE = &H10
Public Const INTERNET_FLAG_NO_AUTH = &H40000
Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000
Public Const INTERNET_FLAG_NO_COOKIES = &H80000
Public Const INTERNET_FLAG_NO_UI = &H200
Public Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100
Public Const INTERNET_FLAG_READ_PREFETCH = &H100000
Public Const INTERNET_FLAG_RESYNCHRONIZE = &H800
Public Const INTERNET_FLAG_SECURE = &H800000
Public Const INTERNET_FLAG_TRANSFER_ASCII = FTP_TRANSFER_TYPE_ASCII 'Transfers the file as ASCII.
Public Const INTERNET_FLAG_TRANSFER_BINARY = FTP_TRANSFER_TYPE_BINARY 'Transfers the file as binary. The following flags determine how the file caching will be done. A combination of the following flags can be used with the transfer type flag. Possible values are described in the following table.

'--- Todo: ADD const from wininet.h
Public Const INTERNET_CONNECTION_MODEM = 1 '---Local system uses a modem to connect to the Internet.
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8 '---No longer used.
Public Const INTERNET_CONNECTION_PROXY = 4 '--- Local system uses a proxy server to connect to the Internet.
Public Const COOKIE_CACHE_ENTRY = &H100000
Public Const TRACK_OFFLINE_CACHE_ENTRY = &H10
Public Const NORMAL_CACHE_ENTRY = &H1
Public Const TRACK_ONLINE_CACHE_ENTRY = &H20
Public Const SPARCE_CACHE_ENTRY = &H10000
Public Const URLHISTORY_CACHE_ENTRY = &H200000
Public Const STICKY_CACHE_ENTRY = &H4

'--- enums
Enum INTERNET_SCHEME
    INTERNET_SCHEME_PARTIAL = -2
    INTERNET_SCHEME_UNKNOWN = -1
    INTERNET_SCHEME_DEFAULT = 0
    INTERNET_SCHEME_FTP
    INTERNET_SCHEME_GOPHER
    INTERNET_SCHEME_HTTP
    INTERNET_SCHEME_HTTPS
    INTERNET_SCHEME_FILE
    INTERNET_SCHEME_NEWS
    INTERNET_SCHEME_MAILTO
    INTERNET_SCHEME_SOCKS
    INTERNET_SCHEME_JAVASCRIPT
    INTERNET_SCHEME_VBSCRIPT
    INTERNET_SCHEME_FIRST = 1
    INTERNET_SCHEME_LAST = INTERNET_SCHEME_VBSCRIPT
End Enum
'--- Types
Type T_FILE_TIME
    lLowDateTime As Long
    lHighDateTime As Long
End Type
Type T_WIN32_FIND_DATA
    lFileAttributes As Long
    ftCreationTime As T_FILE_TIME
    ftLastAccessTime As T_FILE_TIME
    ftLastWriteTime As T_FILE_TIME
    lFileSizeHigh As Long
    lFileSizeLow As Long
    lOID As Long
    sFileName As String
End Type
Type T_INTERNET_CACHE_ENTRY_INFO
    lStructSize As Long
    sSourceUrlName As String
    sLocalFileName As String
    lCacheEntryType As Long
    lUseCount As Long
    lHitRate As Long
    lSizeLow As Long
    lSizeHigh As Long
    ftLastModifiedTime As T_FILE_TIME
    ftExpireTime As T_FILE_TIME
    ftLastAccessTime As T_FILE_TIME
    nHeaderInfo As Integer
    lHeaderInfoSize As Long
    sFileExtension As String
End Type
Type T_URL_COMPONENTS
    lStructSize  As Long
    sScheme As String
    lSchemeLength As Long
    lScheme As INTERNET_SCHEME
    sHostName As String
    lHostNameLength As Long
    lPort As Long
    sUserName As String
    lUserNameLength As Long
    sPassword As String
    lPasswordLength As Long
    sUrlPath As String
    lUrlPathLength As Long
    sExtraInfo As String
    lExtraInfoLength As Long
End Type

'---------------------------------------------------------
'--- INET SECTION - API DECLARE
'---------------------------------------------------------
Private Declare Function CommitUrlCacheEntry Lib "wininet" Alias _
                        "CommitUrlCacheEntryA" _
                        (ByVal lpszUrlName As String, _
                        ByVal lpszLocalFileName As String, _
                        ByRef tftExpireTime As T_FILE_TIME, _
                        ByRef tftLastModifiedTime As T_FILE_TIME, _
                        ByVal lCacheEntryType As Long, _
                        ByVal lpHeaderInfo As Long, _
                        ByVal dwHeaderSize As Long, _
                        ByVal lpszFileExtension As String, _
                        ByVal dwReserved As Long) As Long

Private Declare Function CreateUrlCacheEntry Lib "wininet" Alias _
                        "CreateUrlCacheEntryA" _
                        (ByVal lpszUrlName As String, _
                        ByVal dwExpectedFileSize As Long, _
                        ByVal lpszFileExtension As String, _
                        ByVal lpszFileName As String, _
                        ByVal dwReserved As Long) As Long

Private Declare Function CreateUrlCacheGroup Lib "wininet" _
                        (ByVal dwFlags As Long, _
                         ByVal lpReserved As Long) As Long

Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias _
                        "DeleteUrlCacheEntryA" _
                        (ByVal lpszUrlName As String) As Long
                        
Private Declare Function DeleteUrlCacheGroup Lib "wininet" _
                        (ByVal GroupId As Currency, _
                        ByVal dwFlags As Long, _
                        ByVal lpReserved As Long) As Long

Private Declare Function DllInstall Lib "wininet" _
                        (ByVal bInstall As Long, _
                        ByVal pszCmdLine As String) As Long

Private Declare Function FindCloseUrlCache Lib "wininet" _
                        (ByVal hEnumHandle As Long) As Long

Private Declare Function FindNextUrlCacheGroup Lib "wininet" _
                        (ByVal hFind As Long, _
                        ByRef lpGroupId As Currency, _
                        ByRef plpReserved As Long) As Long

Private Declare Function FtpCommand Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal fExpectResponse As Long, _
                        ByVal dwFlags As Long, _
                        ByVal lpszCommand As String, _
                        ByVal dwContext As Long) As Long
                        
Private Declare Function FtpCreateDirectory Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszDirectory As String) As Long
                        
Private Declare Function FtpDeleteFile Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszFileName As String) As Long

Private Declare Function FtpFindFirstFile Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszSearchFile As String, _
                        ByRef lpFindFileData As T_WIN32_FIND_DATA, _
                        ByVal dwFlags As Long, _
                        ByVal dwContext As Long) As Long

Private Declare Function FtpGetCurrentDirectory Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszCurrentDirectory As String, _
                        ByVal lpdwCurrentDirectory As Long) As Long

Private Declare Function FtpGetFile Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszRemoteFile As String, _
                        ByVal lpszNewFile As String, _
                        ByVal fFailIfExists As Long, _
                        ByVal dwFlagsAndAttributes As Long, _
                        ByVal dwFlags As Long, _
                        ByVal dwContext As Long) As Long

Private Declare Function FtpGetFileSize Lib "wininet" _
                        (ByVal hFile As Long, _
                        ByRef lpdwFileSizeHigh As Long) As Long

Private Declare Function FtpOpenFile Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszFileName As String, _
                        ByVal dwAccess As Long, _
                        ByVal dwFlags As Long, _
                        ByVal dwContext As Long) As Long

Private Declare Function FtpPutFile Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszLocalFile As String, _
                        ByVal lpszNewRemoteFile As String, _
                        ByVal dwFlags As Long, _
                        ByVal dwContext As Long) As Long
 
Private Declare Function FtpRemoveDirectory Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszDirectory As String) As Long

Private Declare Function FtpRenameFile Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszExisting As String, _
                        ByVal lpszNew As String) As Long
                        
Private Declare Function FtpSetCurrentDirectory Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszDirectory As String) As Long

Private Declare Function HttpAddRequestHeaders Lib "wininet" _
                        (ByVal hHttpRequest As Long, _
                        ByVal lpszHeaders As String, _
                        ByVal dwHeadersLength As Long, _
                        ByVal dwModifiers As Long) As Long

Private Declare Function HttpEndRequest Lib "wininet" _
                        (ByVal hRequest As Long, _
                        ByRef lpBuffersOut As Long, _
                        ByVal dwFlags As Long, _
                        ByVal dwContext As Long) As Long

Private Declare Function HttpOpenRequest Lib "wininet" _
                        (ByVal hConnect As Long, _
                        ByVal lpszVerb As String, _
                        ByVal lpszObjectName As String, _
                        ByVal lpszVersion As String, _
                        ByVal lpszReferrer As String, _
                        ByVal lplpszAcceptTypes As String, _
                        ByVal dwFlags As Long, _
                        ByVal dwContext As Long) As Long

Private Declare Function HttpQueryInfo Lib "wininet" _
                        (ByVal hRequest As Long, _
                        ByVal dwInfoLevel As Long, _
                        ByVal lpBuffer As Long, _
                        ByVal lpdwBufferLength As Long, _
                        ByRef lpdwIndex As Long) As Long

Private Declare Function HttpSendRequest Lib "wininet" _
                        (ByVal hRequest As Long, _
                        ByVal lpszHeaders As String, _
                        ByVal dwHeadersLength As Long, _
                        ByVal lpOptional As Long, _
                        ByVal dwOptionalLength As Long) As Long

Private Declare Function InternetAttemptConnect Lib "wininet" _
                        (ByVal dwReserved As Long) As Long

Private Declare Function InternetAutodial Lib "wininet" _
                        (ByVal dwFlags As Long, _
                        ByVal hwndParent As Long) As Long

Private Declare Function InternetAutodialHangup Lib "wininet" _
                        (ByVal dwReserved As Long) As Long

Private Declare Function InternetCanonicalizeUrl Lib "wininet" _
                        (ByVal lpszUrl As String, _
                        ByVal lpszBuffer As String, _
                        ByRef lpdwBufferLength As Long, _
                        ByVal dwFlags As Long) As Long

Private Declare Function InternetCheckConnection Lib "wininet" _
                        (ByVal lpszUrl As String, _
                        ByVal dwFlags As Long, _
                        ByVal dwReserved As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet" _
                        (ByVal hInet As Long) As Long

Private Declare Function InternetCombineUrl Lib "wininet" _
                        (ByVal lpszBaseUrl As String, _
                        ByVal lpszRelativeUrl As String, _
                        ByVal lpszBuffer As String, _
                        ByVal lpdwBufferLength As Long, _
                        ByVal dwFlags As Long) As Long

Private Declare Function InternetConfirmZoneCrossing Lib "wininet" _
                        (ByVal hWnd As Long, _
                        ByVal szUrlPrev As String, _
                        ByVal szUrlNew As String, _
                        ByVal bPost As Long) As Long

Private Declare Function InternetConnect Lib "wininet" _
                        (ByVal hInternet As Long, _
                        ByVal lpszServerName As String, _
                        ByVal lServerPort As Long, _
                        ByVal lpszUserName As String, _
                        ByVal lpszPassword As String, _
                        ByVal dwService As Long, _
                        ByVal dwFlags As Long, _
                        ByVal dwContext As Long) As Long

Private Declare Function InternetCrackUrl Lib "wininet" _
                        (ByVal lpszUrl As String, _
                        ByVal dwUrlLength As Long, _
                        ByVal dwFlags As Long, _
                        ByRef lpUrlComponents As T_URL_COMPONENTS) As Long

Private Declare Function InternetCreateUrl Lib "wininet" _
                        (ByRef lpUrlComponents As T_URL_COMPONENTS, _
                        ByVal dwFlags As Long, _
                        ByVal lpszUrl As String, _
                        ByRef lpdwUrlLength As Long) As Long

Private Declare Function InternetDial Lib "wininet" _
                        (ByVal hwndParent As Long, _
                        ByVal lpszConnectoid As String, _
                        ByVal dwFlags As Long, _
                        ByRef lpdwConnection As Long, _
                        ByVal dwReserved As Long) As Long

Private Declare Function InternetErrorDlg Lib "wininet" _
                        (ByVal hWnd As Long, _
                        ByRef hRequest As Long, _
                        ByVal dwError As Long, _
                        ByVal dwFlags As Long, _
                        ByRef lppvData As Any) As Long
 
Private Declare Function InternetFindNextFile Lib "wininet" _
                        (ByVal hFind As Long, _
                        ByRef lpvFindData As T_WIN32_FIND_DATA) As Long

Private Declare Function InternetGetConnectedState Lib "wininet" _
                        (ByRef lpdwFlags As Long, _
                        ByVal dwReserved As Long) As Long

Private Declare Function InternetGetCookie Lib "wininet" _
                        (ByVal lpszUrl As String, _
                        ByVal lpszCookieName As String, _
                        ByVal lpCookieData As String, _
                        ByRef lpdwSize As Long) As Long
 
Private Declare Function InternetGetLastResponseInfo Lib "wininet" _
                        (ByRef lpdwError As Long, _
                        ByVal lpszBuffer As String, _
                        ByRef lpdwBufferLength As Long) As Long

Private Declare Function InternetGoOnline Lib "wininet" _
                        (ByVal lpszUrl As String, _
                        ByVal hwndParent As Long, _
                        ByVal dwReserved As Long) As Long

Private Declare Function InternetHangUp Lib "wininet" _
                        (ByVal dwConnection As Long, _
                        ByVal dwReserved As Long) As Long

Private Declare Function InternetLockRequestFile Lib "wininet" _
                        (ByVal hInternet As Long, _
                        ByRef lphLockRequestInfo As Long) As Long

Private Declare Function InternetOpen Lib "wininet" Alias _
                        "InternetOpenA" _
                        (ByVal sAgent As String, _
                        ByVal lAccessType As Long, _
                        ByVal sProxyName As String, _
                        ByVal sProxyBypass As String, _
                        ByVal lFlags As Long) As Long
                                    
Private Declare Function InternetOpenUrl Lib "wininet" Alias _
                        "InternetOpenUrlA" _
                        (ByVal hInternetSession As Long, _
                        ByVal lpszUrl As String, _
                        ByVal lpszHeaders As String, _
                        ByVal dwHeadersLength As Long, _
                        ByVal dwFlags As Long, _
                        ByVal dwContext As Long) As Long
                                    
Private Declare Function InternetQueryDataAvailable Lib "wininet" _
                        (ByVal hFile As Long, _
                        ByRef lpdwNumberOfBytesAvailable As Long, _
                        ByVal dwFlags As Long, _
                        ByVal dwContext As Long) As Long

Private Declare Function InternetQueryOption Lib "wininet" _
                        (ByVal hInternet As Long, _
                        ByVal dwOption As Long, _
                        ByRef lpBuffer As Long, _
                        ByRef lpdwBufferLength As Long) As Long

Private Declare Function InternetReadFile Lib "wininet" _
                        (ByVal hFile As Long, _
                        ByVal sBuffer As String, _
                        ByVal lNumBytesToRead As Long, _
                        lNumberOfBytesRead As Long) As Long

Private Declare Function InternetSetCookie Lib "wininet" _
                        (ByVal lpszUrl As String, _
                        ByVal lpszCookieName As String, _
                        ByVal lpszCookieData As Long) As Long

Private Declare Function InternetSetFilePointer Lib "wininet" _
                        (ByVal hFile As Long, _
                        ByVal lDistanceToMove As Long, _
                        ByVal pReserved As Long, _
                        ByVal dwMoveMethod As Long, _
                        ByVal dwContext As Long) As Long

Private Declare Function InternetSetOption Lib "wininet" _
                        (ByVal hInternet As Long, _
                        ByVal dwOption As Long, _
                        ByVal lpBuffer As Long, _
                        ByVal dwBufferLength As Long) As Long

Private Declare Function InternetUnlockRequestFile Lib "wininet" _
                        (ByVal hLockRequestInfo As Long) As Long

Private Declare Function InternetWriteFile Lib "wininet" _
                        (ByVal hFile As Long, _
                        ByRef lpBuffer As Any, _
                        ByVal dwNumberOfBytesToWrite As Long, _
                        ByRef lpdwNumberOfBytesWritten As Long) As Long

Private Declare Function ReadUrlCacheEntryStream Lib "wininet" _
                        (ByVal hUrlCacheStream As Long, _
                        ByVal dwLocation As Long, _
                        ByRef lpBuffer As Long, _
                        ByRef lpdwLen As Long, _
                        ByVal dwReserved As Long) As Long

Private Declare Function RetrieveUrlCacheEntryStream Lib "wininet" _
                        (ByVal lpszUrlName As String, _
                        ByRef lpCacheEntryInfo As T_INTERNET_CACHE_ENTRY_INFO, _
                        ByRef lpdwCacheEntryInfoBufferSize As Long, _
                        ByVal fRandomRead As Long, _
                        ByVal dwReserved As Long) As Long

Private Declare Function RetrieveUrlCacheEntryFile Lib "wininet" _
                        (ByVal lpszUrlName As String, _
                        ByRef lpCacheEntryInfo As T_INTERNET_CACHE_ENTRY_INFO, _
                        ByRef lpdwCacheEntryInfoBufferSize As Long, _
                        ByVal dwReserved As Long) As Long

Private Declare Function SetUrlCacheEntryGroup Lib "wininet" _
                        (ByVal lpszUrlName As String, _
                        ByVal dwFlags As Long, _
                        ByVal GroupId As Currency, _
                        ByVal pbGroupAttributes As Long, _
                        ByVal cbGroupAttributes As Long, _
                        ByVal lpReserved As Long) As Long

Private Declare Function SetUrlCacheEntryInfo Lib "wininet" _
                        (ByVal lpszUrlName As String, _
                        ByRef lpCacheEntryInfo As T_INTERNET_CACHE_ENTRY_INFO, _
                        ByVal dwFieldControl As Long) As Long

Private Declare Function SetUrlCacheGroupAttribute Lib "wininet" _
                        (ByVal gid As Currency, _
                        ByVal dwFlags As Long, _
                        ByVal dwAttributes As Long, _
                        ByRef lpGroupInfo As T_INTERNET_CACHE_ENTRY_INFO, _
                        ByVal lpReserved As Long) As Long

Private Declare Function ShowSecurityInfo Lib "wininet" _
                        () As Long

Private Declare Function UnlockUrlCacheEntryFile Lib "wininet" _
                        (ByVal lpszUrlName As String, _
                        ByVal lpReserved As Long) As Long

Private Declare Function UnlockUrlCacheEntryStream Lib "wininet" _
                        (ByVal hUrlCacheStream As Long, _
                        ByVal dwReserved As Long) As Long

'---------------------------------------------------------
'--- INET SECTION - FUNCTION DECLARE
'---------------------------------------------------------

Public Function netCommitUrlCacheEntry(ByVal sUrlName As String, _
                        ByVal sLocalFileName As String, _
                        ByRef tftExpireTime As T_FILE_TIME, _
                        ByRef tftLastModifiedTime As T_FILE_TIME, _
                        ByVal lCacheEntryType As Long, _
                        ByVal lHeaderInfo As Long, _
                        ByVal lHeaderSize As Long, _
                        ByVal sFileExtension As String, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netCommitUrlCacheEntry"
    
    On Error GoTo EH
    netCommitUrlCacheEntry = _
        CommitUrlCacheEntry(sUrlName, sLocalFileName, tftExpireTime, _
                            tftLastModifiedTime, lCacheEntryType, lHeaderInfo, _
                            lHeaderSize, sFileExtension, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netCreateUrlCacheEntry _
                        (ByVal sUrlName As String, _
                        ByVal lExpectedFileSize As Long, _
                        ByVal sFileExtension As String, _
                        ByVal sFileName As String, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netCreateUrlCacheEntry"
    
    On Error GoTo EH
    netCreateUrlCacheEntry = _
        CreateUrlCacheEntry(sUrlName, lExpectedFileSize, _
                            sFileExtension, sFileName, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netCreateUrlCacheGroup(ByVal lFlags As Long, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netCreateUrlCacheGroup"
    
    On Error GoTo EH
    netCreateUrlCacheGroup = _
        CreateUrlCacheGroup(lFlags, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netDeleteUrlCacheEntry(ByVal sUrlName As String) As Long
    Const FUNC_NAME As String = "netDeleteUrlCacheEntry"
    
    On Error GoTo EH
    netDeleteUrlCacheEntry = _
        DeleteUrlCacheEntry(sUrlName)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netDeleteUrlCacheGroup(ByVal lGroupId As Currency, _
                        ByVal lFlags As Long, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netDeleteUrlCacheGroup"
    
    On Error GoTo EH
    netDeleteUrlCacheGroup = _
        DeleteUrlCacheGroup(lGroupId, lFlags, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netDllInstall(ByVal bInstall As Long, _
                        ByVal sCmdLine As String) As Long
    Const FUNC_NAME As String = "netDllInstall"
    
    On Error GoTo EH
    netDllInstall = _
        DllInstall(bInstall, sCmdLine)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFindCloseUrlCache(ByVal lEnumHandle As Long) As Long
    Const FUNC_NAME As String = "netFindCloseUrlCache"
    
    On Error GoTo EH
    netFindCloseUrlCache = _
        FindCloseUrlCache(lEnumHandle)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFindNextUrlCacheGroup(ByVal lFind As Long, _
                        ByRef lGroupId As Currency, _
                        ByRef lReserved As Long) As Long
    Const FUNC_NAME As String = "netFindNextUrlCacheGroup"
    
    On Error GoTo EH
    netFindNextUrlCacheGroup = _
        FindNextUrlCacheGroup(lFind, lGroupId, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpCommand(ByVal lConnect As Long, _
                        ByVal bExpectResponse As Long, _
                        ByVal lFlags As Long, ByVal sCommand As String, _
                        ByVal lContext As Long) As Long
    Const FUNC_NAME As String = "netFtpCommand"

    On Error GoTo EH
    netFtpCommand = _
        FtpCommand(lConnect, bExpectResponse, lFlags, sCommand, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpCreateDirectory(ByVal lConnect As Long, _
                        ByVal sDirectory As String) As Long
    Const FUNC_NAME As String = "netFtpCreateDirectory"

    On Error GoTo EH
    netFtpCreateDirectory = _
        FtpCreateDirectory(lConnect, sDirectory)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpDeleteFile(ByVal lConnect As Long, _
                        ByVal sFileName As String) As Long
    Const FUNC_NAME As String = "netFtpDeleteFile"

    On Error GoTo EH
    netFtpDeleteFile = _
        FtpDeleteFile(lConnect, sFileName)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpFindFirstFile(ByVal lConnect As Long, _
                        ByVal sSearchFile As String, _
                        ByRef tFindFileData As T_WIN32_FIND_DATA, _
                        ByVal lFlags As Long, _
                        ByVal lContext As Long) As Long
    Const FUNC_NAME As String = "netFtpFindFirstFile"

    On Error GoTo EH
    netFtpFindFirstFile = _
        FtpFindFirstFile(lConnect, sSearchFile, tFindFileData, _
                            lFlags, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpGetCurrentDirectory(ByVal lConnect As Long, _
                        ByVal sCurrentDirectory As String, _
                        ByVal lCurrentDirectory As Long) As Long
    Const FUNC_NAME As String = "netFtpGetCurrentDirectory"

    On Error GoTo EH
    netFtpGetCurrentDirectory = _
        FtpGetCurrentDirectory(lConnect, sCurrentDirectory, lCurrentDirectory)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpGetFile(ByVal lConnect As Long, _
                        ByVal sRemoteFile As String, _
                        ByVal sNewFile As String, _
                        ByVal bFailIfExists As Long, _
                        ByVal lFlagsAndAttributes As Long, _
                        ByVal lFlags As Long, _
                        ByVal lContext As Long) As Long
    Const FUNC_NAME As String = "netFtpGetFile"
    
    On Error GoTo EH
    netFtpGetFile = _
        FtpGetFile(lConnect, sRemoteFile, sNewFile, bFailIfExists, _
                    lFlagsAndAttributes, lFlags, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpGetFileSize(ByVal lFile As Long, _
                        ByRef lFileSizeHigh As Long) As Long
                            
    Const FUNC_NAME As String = "netFtpGetFileSize"

    On Error GoTo EH
    netFtpGetFileSize = _
        FtpGetFileSize(lFile, lFileSizeHigh)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpOpenFile(ByVal lConnect As Long, _
                        ByVal sFileName As String, _
                        ByVal lAccess As Long, _
                        ByVal lFlags As Long, _
                        ByVal lContext As Long) As Long
    Const FUNC_NAME As String = "netFtpOpenFile"

    On Error GoTo EH
    netFtpOpenFile = _
        FtpOpenFile(lConnect, sFileName, lAccess, lFlags, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpPutFile(ByVal lConnect As Long, _
                        ByVal sLocalFile As String, _
                        ByVal sNewRemoteFile As String, _
                        ByVal lFlags As Long, _
                        ByVal lContext As Long) As Long
    Const FUNC_NAME As String = "netFtpPutFile"

    On Error GoTo EH
    netFtpPutFile = _
        FtpPutFile(lConnect, sLocalFile, sNewRemoteFile, lFlags, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpRemoveDirectory(ByVal lConnect As Long, _
                        ByVal sDirectory As String) As Long
    Const FUNC_NAME As String = "netFtpRemoveDirectory"

    On Error GoTo EH
    netFtpRemoveDirectory = _
        FtpRemoveDirectory(lConnect, sDirectory)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpRenameFile(ByVal lConnect As Long, _
                        ByVal sExisting As String, _
                        ByVal sNew As String) As Long

    Const FUNC_NAME As String = "netFtpRenameFile"

    On Error GoTo EH
    netFtpRenameFile = _
        FtpRenameFile(lConnect, sExisting, sNew)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netFtpSetCurrentDirectory(ByVal lConnect As Long, _
                        ByVal sDirectory As String) As Long

    Const FUNC_NAME As String = "netFtpSetCurrentDirectory"

    On Error GoTo EH
    netFtpSetCurrentDirectory = _
        FtpSetCurrentDirectory(lConnect, sDirectory)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netHttpAddRequestHeaders(ByVal lHttpRequest As Long, _
                        ByVal sHeaders As String, _
                        ByVal lHeadersLength As Long, _
                        ByVal lModifiers As Long) As Long
    Const FUNC_NAME As String = "netHttpAddRequestHeaders"

    On Error GoTo EH
    netHttpAddRequestHeaders = _
        HttpAddRequestHeaders(lHttpRequest, sHeaders, lHeadersLength, _
                                lModifiers)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netHttpEndRequest(ByVal lRequest As Long, _
                        ByRef lBuffersOut As Long, _
                        ByVal lFlags As Long, _
                        ByVal lContext As Long) As Long
    Const FUNC_NAME As String = "netHttpEndRequest"

    On Error GoTo EH
    netHttpEndRequest = _
        HttpEndRequest(lRequest, lBuffersOut, lFlags, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netHttpOpenRequest(ByVal lConnect As Long, _
                        ByVal sVerb As String, _
                        ByVal sObjectName As String, _
                        ByVal sVersion As String, _
                        ByVal sReferrer As String, _
                        ByVal sAcceptTypes As String, _
                        ByVal lFlags As Long, _
                        ByVal lContext As Long) As Long

    Const FUNC_NAME As String = "netHttpEndRequest"

    On Error GoTo EH
    netHttpOpenRequest = _
        HttpOpenRequest(lConnect, sVerb, sObjectName, sVersion, _
                        sReferrer, sAcceptTypes, lFlags, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netHttpQueryInfo(ByVal lRequest As Long, _
                        ByVal lInfoLevel As Long, _
                        ByVal lBuffer As Long, _
                        ByVal lBufferLength As Long, _
                        ByRef lIndex As Long) As Long
    
    Const FUNC_NAME As String = "netHttpEndRequest"

    On Error GoTo EH
    netHttpQueryInfo = _
        HttpQueryInfo(lRequest, lInfoLevel, lBuffer, lBufferLength, lIndex)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netHttpSendRequest(ByVal lRequest As Long, _
                        ByVal sHeaders As String, _
                        ByVal lHeadersLength As Long, _
                        ByVal lOptional As Long, _
                        ByVal lOptionalLength As Long) As Long
                        
    Const FUNC_NAME As String = "netHttpSendRequest"

    On Error GoTo EH
    netHttpSendRequest = _
        HttpSendRequest(lRequest, sHeaders, lHeadersLength, lOptional, _
                        lOptionalLength)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetAttemptConnect(ByVal lReserved As Long) As Long
                        
    Const FUNC_NAME As String = "netInternetAttemptConnect"

    On Error GoTo EH
    netInternetAttemptConnect = _
        InternetAttemptConnect(lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetAutodial(ByVal lFlags As Long, _
                        ByVal hwndParent As Long) As Long
    Const FUNC_NAME As String = "netInternetAttemptConnect"

    On Error GoTo EH
    netInternetAutodial = _
        InternetAutodial(lFlags, hwndParent)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetAutodialHangup(ByVal lReserved As Long) As Long
                        
    Const FUNC_NAME As String = "netInternetAutodialHangup"

    On Error GoTo EH
    netInternetAutodialHangup = _
        InternetAutodialHangup(lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetCanonicalizeUrl(ByVal sUrl As String, _
                        ByVal sBuffer As String, _
                        ByRef lBufferLength As Long, _
                        ByVal lFlags As Long) As Long
                        
    Const FUNC_NAME As String = "netInternetCanonicalizeUrl"

    On Error GoTo EH
    netInternetCanonicalizeUrl = _
        InternetCanonicalizeUrl(sUrl, sBuffer, lBufferLength, lFlags)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetCheckConnection(ByVal sUrl As String, _
                        ByVal lFlags As Long, _
                        ByVal lReserved As Long) As Long
                        
    Const FUNC_NAME As String = "netInternetCheckConnection"

    On Error GoTo EH
    netInternetCheckConnection = _
        InternetCheckConnection(sUrl, lFlags, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetCloseHandle(ByVal hInet As Long) As Long
    Const FUNC_NAME As String = "netIntenetOpen"

    On Error GoTo EH
    netInternetCloseHandle = _
        InternetCloseHandle(hInet)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetCombineUrl(ByVal sBaseUrl As String, _
                        ByVal sRelativeUrl As String, _
                        ByVal sBuffer As String, _
                        ByVal lBufferLength As Long, _
                        ByVal lFlags As Long) As Long
                        
    Const FUNC_NAME As String = "netInternetCombineUrl"

    On Error GoTo EH
    netInternetCombineUrl = _
        InternetCombineUrl(sBaseUrl, sRelativeUrl, sBuffer, lBufferLength, _
                            lFlags)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetConfirmZoneCrossing(ByVal hWnd As Long, _
                        ByVal sUrlPrev As String, _
                        ByVal sUrlNew As String, _
                        ByVal bPost As Long) As Long
                        
    Const FUNC_NAME As String = "netInternetConfirmZoneCrossing"

    On Error GoTo EH
    netInternetConfirmZoneCrossing = _
        InternetConfirmZoneCrossing(hWnd, sUrlPrev, sUrlNew, bPost)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetConnect(ByVal hInternet As Long, _
                        ByVal sServerName As String, _
                        ByVal lServerPort As Long, _
                        ByVal sUserName As String, _
                        ByVal sPassword As String, _
                        ByVal lService As Long, _
                        ByVal lFlags As Long, _
                        ByVal lContext As Long) As Long
                        
    Const FUNC_NAME As String = "netInternetConnect"

    On Error GoTo EH
    netInternetConnect = _
        InternetConnect(hInternet, sServerName, lServerPort, sUserName, _
                        sPassword, lService, lFlags, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetCrackUrl(ByVal sUrl As String, _
                        ByVal lUrlLength As Long, _
                        ByVal lFlags As Long, _
                        ByRef tucUrlComponents As T_URL_COMPONENTS) As Long
                        
    Const FUNC_NAME As String = "netInternetCrackUrl"

    On Error GoTo EH
    netInternetCrackUrl = _
        InternetCrackUrl(sUrl, lUrlLength, lFlags, tucUrlComponents)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetCreateUrl(ByRef tucUrlComponents As T_URL_COMPONENTS, _
                        ByVal lFlags As Long, _
                        ByVal sUrl As String, _
                        ByRef lUrlLength As Long) As Long
                        
    Const FUNC_NAME As String = "netInternetCreateUrl"

    On Error GoTo EH
    netInternetCreateUrl = _
        InternetCreateUrl(tucUrlComponents, lFlags, sUrl, lUrlLength)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetDial(ByVal hwndParent As Long, _
                        ByVal sConnectoid As String, _
                        ByVal lFlags As Long, _
                        ByRef lConnection As Long, _
                        ByVal lReserved As Long) As Long
                        
    Const FUNC_NAME As String = "netInternetDial"

    On Error GoTo EH
    netInternetDial = _
        InternetDial(hwndParent, sConnectoid, lFlags, lConnection, lReserved)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetErrorDlg(ByVal hWnd As Long, _
                        ByRef hRequest As Long, _
                        ByVal lError As Long, _
                        ByVal lFlags As Long, _
                        ByRef lppvData As Long) As Long
 
    Const FUNC_NAME As String = "netInternetErrorDlg"

    On Error GoTo EH
    netInternetErrorDlg = _
        InternetErrorDlg(hWnd, hRequest, lError, lFlags, lppvData)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetFindNextFile(ByVal hFind As Long, _
                        ByRef lpvFindData As T_WIN32_FIND_DATA) As Long
    
    Const FUNC_NAME As String = "netInternetFindNextFile"

    On Error GoTo EH
    netInternetFindNextFile = _
        InternetFindNextFile(hFind, lpvFindData)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetGetConnectedState(ByRef lFlags As Long, _
                        ByVal lReserved As Long) As Long
    
    Const FUNC_NAME As String = "netInternetGetConnectedState"

    On Error GoTo EH
    netInternetGetConnectedState = _
        InternetGetConnectedState(lFlags, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetGetCookie(ByVal sUrl As String, _
                        ByVal sCookieName As String, _
                        ByVal sCookieData As String, _
                        ByRef lSize As Long) As Long
    
    Const FUNC_NAME As String = "netInternetGetCookie"

    On Error GoTo EH
    netInternetGetCookie = _
        InternetGetCookie(sUrl, sCookieName, sCookieData, lSize)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetGetLastResponseInfo _
                        (ByRef lError As Long, _
                        ByVal sBuffer As String, _
                        ByRef lBufferLength As Long) As Long
    
    Const FUNC_NAME As String = "netInternetGetLastResponseInfo"

    On Error GoTo EH
    netInternetGetLastResponseInfo = _
        InternetGetLastResponseInfo(lError, sBuffer, lBufferLength)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetGoOnline(ByVal sUrl As String, _
                        ByVal hwndParent As Long, _
                        ByVal lReserved As Long) As Long
    
    Const FUNC_NAME As String = "netInternetGoOnline"

    On Error GoTo EH
    netInternetGoOnline = _
        InternetGoOnline(sUrl, hwndParent, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetHangUp(ByVal lConnection As Long, _
                        ByVal lReserved As Long) As Long
    
    Const FUNC_NAME As String = "netInternetGoOnline"

    On Error GoTo EH
    netInternetHangUp = _
        InternetHangUp(lConnection, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetLockRequestFile(ByVal hInternet As Long, _
                        ByRef lphLockRequestInfo As Long) As Long
    
    Const FUNC_NAME As String = "netInternetLockRequestFile"

    On Error GoTo EH
    netInternetLockRequestFile = _
        InternetLockRequestFile(hInternet, lphLockRequestInfo)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetOpen(ByVal sAgent As String, _
                        ByVal lAccessType As Long, _
                        ByVal sProxyName As String, _
                        ByVal sProxyBypass As String, _
                        ByVal lFlags As Long) As Long
    Const FUNC_NAME As String = "netIntenetOpen"

    On Error GoTo EH
    netInternetOpen = _
        InternetOpen(sAgent, lAccessType, sProxyName, sProxyBypass, lFlags)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetOpenUrl(ByVal lInternetSession As Long, _
                        ByVal sUrl As String, _
                        ByVal sHeaders As String, _
                        ByVal lHeadersLength As Long, _
                        ByVal lFlags As Long, _
                        ByVal lContext As Long) As Long
    Const FUNC_NAME As String = "netInternetOpenUrl"

    On Error GoTo EH
    netInternetOpenUrl = _
        InternetOpenUrl(lInternetSession, sUrl, sHeaders, lHeadersLength, _
                        lFlags, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetQueryDataAvailable(ByVal lFile As Long, _
                        ByRef lNumberOfBytesAvailable As Long, _
                        ByVal lFlags As Long, _
                        ByVal lContext As Long) As Long
    Const FUNC_NAME As String = "netInternetQueryDataAvailable"

    On Error GoTo EH
    netInternetQueryDataAvailable = _
        InternetQueryDataAvailable(lFile, lNumberOfBytesAvailable, _
                                    lFlags, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetQueryOption(ByVal hInternet As Long, _
                        ByVal lOption As Long, _
                        ByRef lBuffer As Long, _
                        ByRef lpdwBufferLength As Long) As Long
    Const FUNC_NAME As String = "netInternetQueryDataAvailable"

    On Error GoTo EH
    netInternetQueryOption = _
        InternetQueryOption(hInternet, lOption, lBuffer, lpdwBufferLength)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetReadFile(ByVal lFile As Long, _
                        ByRef sBuffer As String, _
                        ByVal lNumBytesToRead As Long, _
                        lNumberOfBytesRead As Long) As Long
    Const FUNC_NAME As String = "netInternetReadFile"

    On Error GoTo EH
    netInternetReadFile = _
        InternetReadFile(lFile, sBuffer, lNumBytesToRead, lNumberOfBytesRead)

    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetSetCookie(ByVal sUrl As String, _
                        ByVal sCookieName As String, _
                        ByVal lCookieData As Long) As Long
    Const FUNC_NAME As String = "netInternetSetCookie"

    On Error GoTo EH
    netInternetSetCookie = _
        InternetSetCookie(sUrl, sCookieName, lCookieData)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetSetFilePointer(ByVal hFile As Long, _
                        ByVal lDistanceToMove As Long, _
                        ByVal pReserved As Long, _
                        ByVal lMoveMethod As Long, _
                        ByVal lContext As Long) As Long
    Const FUNC_NAME As String = "netInternetSetCookie"

    On Error GoTo EH
    netInternetSetFilePointer = _
        InternetSetFilePointer(hFile, lDistanceToMove, pReserved = 0, _
                            lMoveMethod, lContext)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetSetOption(ByVal hInternet As Long, _
                        ByVal lOption As Long, _
                        ByVal lBuffer As Long, _
                        ByVal lBufferLength As Long) As Long
    Const FUNC_NAME As String = "netInternetSetCookie"

    On Error GoTo EH
    netInternetSetOption = _
        InternetSetOption(hInternet, lOption, lBuffer, lBufferLength)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetUnlockRequestFile _
                        (ByVal hLockRequestInfo As Long) As Long
    Const FUNC_NAME As String = "netInternetUnlockRequestFile"

    On Error GoTo EH
    netInternetUnlockRequestFile = _
        InternetUnlockRequestFile(hLockRequestInfo)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netInternetWriteFile(ByVal hFile As Long, _
                        ByRef lBuffer As Long, _
                        ByVal lNumberOfBytesToWrite As Long, _
                        ByRef lNumberOfBytesWritten As Long) As Long
    Const FUNC_NAME As String = "netInternetWriteFile"

    On Error GoTo EH
    netInternetWriteFile = _
        InternetWriteFile(hFile, lBuffer, _
            lNumberOfBytesToWrite, lNumberOfBytesWritten)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netReadUrlCacheEntryStream(ByVal hUrlCacheStream As Long, _
                        ByVal lLocation As Long, _
                        ByRef lBuffer As Long, _
                        ByRef lLen As Long, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netReadUrlCacheEntryStream"

    On Error GoTo EH
    netReadUrlCacheEntryStream = _
        ReadUrlCacheEntryStream(hUrlCacheStream, lLocation, _
            lBuffer, lLen, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netRetrieveUrlCacheEntryStream(ByVal sUrlName As String, _
                        ByRef vCacheEntryInfo As T_INTERNET_CACHE_ENTRY_INFO, _
                        ByRef lCacheEntryInfoBufferSize As Long, _
                        ByVal bRandomRead As Long, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netRetrieveUrlCacheEntryStream"

    On Error GoTo EH
    netRetrieveUrlCacheEntryStream = _
        RetrieveUrlCacheEntryStream(sUrlName, vCacheEntryInfo, _
            lCacheEntryInfoBufferSize, bRandomRead, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netRetrieveUrlCacheEntryFile(ByVal sUrlName As String, _
                        ByRef vCacheEntryInfo As T_INTERNET_CACHE_ENTRY_INFO, _
                        ByRef lCacheEntryInfoBufferSize As Long, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netRetrieveUrlCacheEntryFile"

    On Error GoTo EH
    netRetrieveUrlCacheEntryFile = _
        RetrieveUrlCacheEntryFile(sUrlName, vCacheEntryInfo, _
            lCacheEntryInfoBufferSize, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netSetUrlCacheEntryGroup(ByVal sUrlName As String, _
                        ByVal lFlags As Long, _
                        ByVal GroupId As Currency, _
                        ByVal pbGroupAttributes As Long, _
                        ByVal cbGroupAttributes As Long, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netSetUrlCacheEntryGroup"

    On Error GoTo EH
    netSetUrlCacheEntryGroup = _
        SetUrlCacheEntryGroup(sUrlName, lFlags, GroupId, _
            pbGroupAttributes = 0, cbGroupAttributes = 0, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netSetUrlCacheEntryInfo(ByVal sUrlName As String, _
                        ByRef vCacheEntryInfo As T_INTERNET_CACHE_ENTRY_INFO, _
                        ByVal lFieldControl As Long) As Long
    Const FUNC_NAME As String = "netSetUrlCacheEntryInfo"

    On Error GoTo EH
    netSetUrlCacheEntryInfo = _
        SetUrlCacheEntryInfo(sUrlName, vCacheEntryInfo, _
            lFieldControl)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netSetUrlCacheGroupAttribute(ByVal gid As Currency, _
                        ByVal lFlags As Long, _
                        ByVal lAttributes As Long, _
                        ByRef lpGroupInfo As T_INTERNET_CACHE_ENTRY_INFO, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "SetUrlCacheGroupAttribute"

    On Error GoTo EH
    netSetUrlCacheGroupAttribute = _
        SetUrlCacheGroupAttribute(gid, lFlags, lAttributes, lpGroupInfo, _
            lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netShowSecurityInfo() As Long
    Const FUNC_NAME As String = "netShowSecurityInfo"

    On Error GoTo EH
    netShowSecurityInfo = ShowSecurityInfo()
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netUnlockUrlCacheEntryFile(ByVal sUrlName As String, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netUnlockUrlCacheEntryFile"

    On Error GoTo EH
    netUnlockUrlCacheEntryFile = _
        UnlockUrlCacheEntryFile(sUrlName, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function netUnlockUrlCacheEntryStream(ByVal hUrlCacheStream As Long, _
                        ByVal lReserved As Long) As Long
    Const FUNC_NAME As String = "netUnlockUrlCacheEntryFile"

    On Error GoTo EH
    netUnlockUrlCacheEntryStream = _
        UnlockUrlCacheEntryStream(hUrlCacheStream, lReserved = 0)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function
