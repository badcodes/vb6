Attribute VB_Name = "modReturnfunc"
Option Explicit

Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Const PAGE_EXECUTE_READWRITE& = &H40&

Public Function IInternetProtocol_Read(ByVal This As olelib2.IInternetProtocol, ByVal pv As Long, ByVal cb As Long, pcbRead As Long) As Long

    Dim oPH As clsZipHandler
   
    Set oPH = This
      
    IInternetProtocol_Read = oPH.Read(pv, cb, pcbRead)
   
End Function

Public Function ReplaceVTableEntry(ByVal oObject As Long, ByVal nEntry As Integer, ByVal pFunc As Long) As Long

    Dim pFuncOld As Long, pVTableHead As Long
    Dim pFuncTmp As Long, lOldProtect As Long
     
    MoveMemory pVTableHead, ByVal oObject, 4
    pFuncTmp = pVTableHead + (nEntry - 1) * 4
    MoveMemory pFuncOld, ByVal pFuncTmp, 4
    
    If pFuncOld <> pFunc Then
        
        VirtualProtect pFuncTmp, 4, PAGE_EXECUTE_READWRITE, lOldProtect
        MoveMemory ByVal pFuncTmp, pFunc, 4
        VirtualProtect pFuncTmp, 4, lOldProtect, lOldProtect
        
    End If
    
    ReplaceVTableEntry = pFuncOld
    
End Function

Public Function IInternetProtocol_Start(ByVal This As olelib2.IInternetProtocol, ByVal szURL As Long, ByVal pOIProtSink As olelib2.IInternetProtocolSink, ByVal pOIBindInfo As olelib.IInternetBindInfo, ByVal grfPI As olelib.PI_FLAGS, dwReserved As olelib.PROTOCOLFILTERDATA) As Long

    Dim oPH As clsZipHandler
   
    Set oPH = This
    IInternetProtocol_Start = oPH.Start(szURL, pOIProtSink, pOIBindInfo, grfPI, dwReserved)

End Function

Public Function IInternetProtocolInfo_ParseUrl(ByVal This As olelib.IInternetProtocolInfo, ByVal pwzUrl As Long, ByVal PARSEACTION As olelib.PARSEACTION, ByVal dwParseFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long) As HRESULTS

    Dim oPH As clsZipHandler
    Set oPH = This

    'MsgBox "modReturnfunc::IInternetProtocolInfo_ParseUrl Begin"

    IInternetProtocolInfo_ParseUrl = oPH.ParseUrl(pwzUrl, PARSEACTION, dwParseFlags, pwzResult, cchResult, pcchResult, dwReserved)

End Function

Public Function IInternetProtocolInfo_QueryInfo(ByVal This As olelib.IInternetProtocolInfo, ByVal pwzUrl As Long, ByVal OueryOption As olelib.QueryOption, ByVal dwQueryFlags As Long, ByVal pBuffer As Long, ByVal cbBuffer As Long, pcbBuf As Long, ByVal dwReserved As Long) As HRESULTS
    Dim oPH As clsZipHandler
    Set oPH = This
    IInternetProtocolInfo_QueryInfo = oPH.QueryInfo(pwzUrl, OueryOption, dwQueryFlags, pBuffer, cbBuffer, pcbBuf, dwReserved)
End Function

Public Function IInternetProtocolInfo_CombineUrl(ByVal This As olelib.IInternetProtocolInfo, ByVal pwzBaseUrl As Long, ByVal pwzRelativeUrl As Long, ByVal dwCombineFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long) As HRESULTS
    Dim oPH As clsZipHandler
    Set oPH = This
    IInternetProtocolInfo_CombineUrl = oPH.CombineUrl(pwzBaseUrl, pwzRelativeUrl, dwCombineFlags, pwzResult, cchResult, pcchResult, dwReserved)

End Function

Public Function MimeType(ByVal sUrl As String) As String
    sUrl = LiNVBLib.RightRight(sUrl, ".", vbBinaryCompare, ReturnEmptyStr)
    sUrl = LiNVBLib.LeftLeft(sUrl, "#", vbBinaryCompare, ReturnOriginalStr)
    sUrl = LiNVBLib.LeftLeft(sUrl, "?", vbBinaryCompare, ReturnOriginalStr)
    sUrl = LCase$(sUrl)
    Select Case sUrl
        Case "jpg", "jpeg"
            MimeType = "image/jpeg"
        Case "gif"
            MimeType = "image/gif"
        Case "htm", "html"
            MimeType = "text/html"
        Case "zip"
            MimeType = "application/zip"
        Case "mp3"
            MimeType = "audio/mpeg"
        Case "m3u", "pls", "xpl"
            MimeType = "audio/x-mpegurl"
        Case "txt", "text"
            MimeType = "text/plain"
        Case "css"
            MimeType = "text/css"
        Case Else
            MimeType = "*/*"
    End Select
End Function
Public Function CanonicalizeUrl(ByVal sUrl As String, dwFlags As Long)
Dim sResult As String
Dim lbuf As Long
lbuf = 512
sResult = String(lbuf, 0)
olelib.UrlCanonicalizeA ByVal sUrl, ByVal sResult, lbuf, dwFlags
sResult = LeftLeft(sResult, vbNullChar, vbBinaryCompare, ReturnOriginalStr)
CanonicalizeUrl = sResult
End Function
