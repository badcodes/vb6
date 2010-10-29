VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUTF8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
'   cUTF8
'       Lots of utility functions for deteremining which codepage and LCID
'       to use for an arbitrary string, very helpful adjunct in multinational
'       applications where you may not know the language.
'
'   Author: MichKa
'
'   Copyright ?1999 Trigeminal Software, Inc. All Rights Reserved.
'
'   You may use this code freely in your applications, you just can't
'   put it in a book or article and call it your own.
'-----------------------------------------------------------------------------
' Simple UTF-8 Encoder/Decoder using API

' Use character sets such as Japanese on the Forms 2.0 controls as only these are unicode enabled

' The actual functions used from the API's are MultiByteToWideChar and WideCharToMultiByte

' For a complete list of Code Pages you can use look in the Module basInCodePage !
' Example... If you want to use UTF-7 Encoding/Decoding you would set sCodePage = CP_UTF7

Option Compare Binary
Option Explicit

Private Const CP_ACP = 0
Private Const CP_NONE = 0
Private Const CP_OEMCP = 1
Private Const CP_WINANSI = 1004
Private Const CP_WINUNICODE = 1200
Private Const MB_PRECOMPOSED = &H1
Private Const MB_COMPOSITE = &H2
Private Const MB_USEGLYPHCHARS = &H4

' //Arabic
Private Const CP_AWIN = 101 ' //Bidi Windows codepage
Private Const CP_709 = 102  ' //MS-DOS Arabic Support CP 709
Private Const CP_720 = 103  ' //MS-DOS Arabic Support CP 720
Private Const CP_A708 = 104 ' //ASMO 708
Private Const CP_A449 = 105 ' //ASMO 449+
Private Const CP_TARB = 106 ' //MS Transparent Arabic
Private Const CP_NAE = 107  ' //Nafitha Enhanced Arabic Char Set
Private Const CP_V4 = 108   ' //Nafitha v 4.0
Private Const CP_MA2 = 109  ' //Mussaed Al Arabi (MA/2) CP 786
Private Const CP_I864 = 110 ' //IBM Arabic Supplement CP 864
Private Const CP_A437 = 111 ' //Ansi 437 codepage
Private Const CP_AMAC = 112 ' //Macintosh Cod Page

' //Hebrew
Private Const CP_HWIN = 201 ' //Bidi Windows codepage
Private Const CP_862I = 202 ' //IBM Hebrew Supplement CP 862
Private Const CP_7BIT = 203 ' //IBM Hebrew Supplement CP 862 Folded
Private Const CP_ISO = 204   ' //ISO Hebrew 8859-8 Character Set
Private Const CP_H437 = 205 ' //Ansi 437 codepage
Private Const CP_HMAC = 206 ' //Macintosh Cod Page

' /*************************************************************************
'    Code Pages
' *************************************************************************/
'Private Const CP_OEM_437 = 437
'Private Const CP_ARABICDOS = 708
'Private Const CP_DOS720 = 720
'Private Const CP_IBM850 = 850
'Private Const CP_IBM852 = 852
'Private Const CP_DOS862 = 862
'Private Const CP_IBM866 = 866
'Private Const CP_THAI = 874
'Private Const CP_JAPAN = 932
'Private Const CP_CHINA = 936
'Private Const CP_KOREA = 949
'Private Const CP_TAIWAN = 950
'Private Const CP_EASTEUROPE = 1250
'Private Const CP_RUSSIAN = 1251
'Private Const CP_WESTEUROPE = 1252
'Private Const CP_GREEK = 1253
'Private Const CP_TURKISH = 1254
'Private Const CP_HEBREW = 1255
'Private Const CP_ARABIC = 1256
'Private Const CP_BALTIC = 1257
'Private Const CP_VIETNAMESE = 1258
'Private Const CP_ASCII = 20127
'Private Const CP_RUSSIANKOI8R = 20866
'Private Const CP_RUSSIANKOI8U = 21866
'Private Const CP_ISOLATIN1 = 28591
'Private Const CP_ISOEASTEUROPE = 28592
'Private Const CP_ISOTURKISH = 28593
'Private Const CP_ISOBALTIC = 28594
'Private Const CP_ISORUSSIAN = 28595
'Private Const CP_ISOARABIC = 28596
'Private Const CP_ISOGREEK = 28597
'Private Const CP_ISOHEBREW = 28598
'Private Const CP_ISOTURKISH2 = 28599
'Private Const CP_ISOLATIN9 = 28605
'Private Const CP_HEBREWLOG = 38598
'Private Const CP_USER = 50000
'Private Const CP_AUTOALL = 50001
'Private Const CP_JAPANNHK = 50220
'Private Const CP_JAPANESC = 50221
'Private Const CP_JAPANSIO = 50222
'Private Const CP_KOREAISO = 50225
'Private Const CP_TAIWANISO = 50227
'Private Const CP_CHINAISO = 50229
'Private Const CP_AUTOJAPAN = 50932
'Private Const CP_AUTOCHINA = 50936
'Private Const CP_AUTOKOREA = 50949
'Private Const CP_AUTOTAIWAN = 50950
'Private Const CP_AUTORUSSIAN = 51251
'Private Const CP_AUTOGREEK = 51253
'Private Const CP_AUTOARABIC = 51256
'Private Const CP_JAPANEUC = 51932
'Private Const CP_CHINAEUC = 51936
'Private Const CP_KOREAEUC = 51949
'Private Const CP_TAIWANEUC = 51950
'Private Const CP_CHINAHZ = 52936
'Private Const CP_MAC_ROMAN = 10000
'Private Const CP_MAC_JAPAN = 10001
'Private Const CP_MAC_ARABIC = 10004
'Private Const CP_MAC_GREEK = 10006
'Private Const CP_MAC_CYRILLIC = 10007
'Private Const CP_MAC_LATIN2 = 10029
'Private Const CP_MAC_TURKISH = 10081
Private Const CP_DEFAULT = CP_ACP

Private Const CP_JOHAB = 1361
Private Const CP_SYMBOL = 42
Private Const CP_UTF8 = 65001
Private Const CP_UTF7 = 65000
Private Const CP_UNICODELITTLE = 1200
Private Const CP_UNICODEBIG = 1201
Private Const CP_UNKNOWN = -1

Private Const MB_ERR_INVALID_CHARS = &H8             ' /* error for invalid chars */

Private Const WC_DEFAULTCHECK = &H100                ' /* check for default char */
Private Const WC_COMPOSITECHECK = &H200              ' /* convert composite to precomposed */
Private Const WC_DISCARDNS = &H10                    ' /* discard non-spacing chars */
Private Const WC_SEPCHARS = &H20                     ' /* generate separate chars */
Private Const WC_DEFAULTCHAR = &H40                  ' /* replace w/ default char */

Private Type FONTSIGNATURE
        fsUsb(4) As Long
        fsCsb(2) As Long
End Type

Private Type CHARSETINFO
        ciCharset As Long
        ciACP As Long
        fs As FONTSIGNATURE
End Type

Private Const LOCALE_IDEFAULTCODEPAGE = &HB
Private Const LOCALE_IDEFAULTANSICODEPAGE = &H1004
Private Const TCI_SRCCODEPAGE = 2

Private Declare Function GetACP Lib "Kernel32" () As Long
Private Declare Function GetLocaleInfoA Lib "Kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "Kernel32" () As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function TranslateCharsetInfo Lib "gdi32" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long

' The OS functions, if you prefer to use them
Private Declare Function MultiByteToWideChar Lib "Kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "Kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long



'--------------------------------
'   AToW
'
'   ANSI to UNICODE conversion, via a given codepage.
'--------------------------------
Public Function AToW(ByVal st As String, Optional ByVal cpg As Long = -1, Optional lFlags As Long = 0) As String
    Dim stBuffer As String
    Dim cwch As Long
    Dim pwz As Long
    Dim pwzBuffer As Long
        
    If cpg = -1 Then cpg = GetACP()
    pwz = StrPtr(st)
    cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, 0&, 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer))
    AToW = Left$(stBuffer, cwch - 1)
End Function

'--------------------------------
'   WToA
'
'   UNICODE to ANSI conversion, via a given codepage
'--------------------------------
Public Function WToA(ByVal st As String, Optional ByVal cpg As Long = -1, Optional lFlags As Long = 0) As String
    Dim stBuffer As String
    Dim cwch As Long
    Dim pwz As Long
    Dim pwzBuffer As Long
    Dim lpUsedDefaultChar As Long
    
    If cpg = -1 Then cpg = GetACP()
    pwz = StrPtr(st)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
    WToA = Left$(stBuffer, cwch - 1)
End Function

'----------------------------------------------------------------------------------------
'   ChsFromLocale
'
'   The OS version
'----------------------------------------------------------------------------------------
Public Function ChsFromLocale(lcid As Long) As Long
    Dim cwc As Long
    Dim cpg As Long
    Dim stBuffer As String
    Dim cs As CHARSETINFO
    
    stBuffer = String$(10, vbNullChar)
    cwc = GetLocaleInfoA(lcid, LOCALE_IDEFAULTANSICODEPAGE, _
     stBuffer, Len(stBuffer))

    If cwc > 0 Then
        cpg = Val(Left$(stBuffer, cwc - 1))
        
        If TranslateCharsetInfo(ByVal cpg, cs, _
         TCI_SRCCODEPAGE) Then
            ChsFromLocale = cs.ciCharset
        End If
    End If
End Function

'%E6%B5%8B%E8%AF%95 =测试
'这比较quick & dirty，但起码能运行，呵呵
Public Sub DecodeURL(ByVal strText As String, ByRef bytContent() As Byte)
    Dim o_intItems As Integer
    Dim o_strChars() As String
    
    If Left(strText, 1) = "%" Then strText = Right(strText, Len(strText) - 1)
        
    o_strChars = Split(strText, "%")
    
    ReDim bytContent(0 To UBound(o_strChars))
    
    For o_intItems = 0 To UBound(o_strChars)
        bytContent(o_intItems) = Val("&H" & o_strChars(o_intItems))
    Next
End Sub

Public Function EncodeUTF8(ByVal cnvUni As String)
    If cnvUni <> "" Then
        EncodeUTF8 = StrConv(WToA(cnvUni, CP_UTF8, 0), vbUnicode)
    End If
End Function

Public Function DecodeUTF8(ByVal cnvUni As String)
    If cnvUni <> "" Then
        Dim cnvUni2 As String
        cnvUni2 = WToA(cnvUni, CP_ACP)
        DecodeUTF8 = AToW(cnvUni2, CP_UTF8)
    End If
End Function

               

