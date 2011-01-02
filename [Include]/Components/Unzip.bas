Attribute VB_Name = "MUnzip"
Option Explicit


'-- Please Do Not Remove These Comment Lines!
'----------------------------------------------------------------
'-- Sample VB 5 / VB 6 code to drive unzip32.dll
'-- Contributed to the Info-ZIP project by Mike Le Voi
'--
'-- Contact me at: mlevoi@modemss.brisnet.org.au
'--
'-- Visit my home page at: http://modemss.brisnet.org.au/~mlevoi
'--
'-- Use this code at your own risk. Nothing implied or warranted
'-- to work on your machine :-)
'----------------------------------------------------------------
'--
'-- This Source Code Is Freely Available From The Info-ZIP Project
'-- Web Server At:
'-- ftp://ftp.info-zip.org/pub/infozip/infozip.html
'--
'-- A Very Special Thanks To Mr. Mike Le Voi
'-- And Mr. Mike White
'-- And The Fine People Of The Info-ZIP Group
'-- For Letting Me Use And Modify Their Original
'-- Visual Basic 5.0 Code! Thank You Mike Le Voi.
'-- For Your Hard Work In Helping Me Get This To Work!!!
'---------------------------------------------------------------
'--
'-- Contributed To The Info-ZIP Project By Raymond L. King.
'-- Modified June 21, 1998
'-- By Raymond L. King
'-- Custom Software Designers
'--
'-- Contact Me At: king@ntplx.net
'-- ICQ 434355
'-- Or Visit Our Home Page At: http://www.ntplx.net/~king
'--
'---------------------------------------------------------------
'--
'-- Modified August 17, 1998
'-- by Christian Spieler
'-- (implemented sort of a "real" user interface)
'-- Modified May 11, 2003
'-- by Christian Spieler
'-- (use late binding for referencing the common dialog)
'--
'---------------------------------------------------------------

'-- This Assumes UNZIP32.DLL Is In Your \Windows\System Directory!
Public Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" _
  (ByVal ifnc As Long, ByRef ifnv As UNZIPnames, _
   ByVal xfnc As Long, ByRef xfnv As UNZIPnames, _
   dcll As DCLIST, Userf As USERFUNCTION) As unzReturnCode

Public Declare Sub UzpVersion2 Lib "unzip32.dll" (uzpv As UZPVER)
Public Declare Function Wiz_Validate Lib "unzip32.dll" (archive As String, AllCodes As Long) As unzReturnCode


'-- C Style argv
Public Type UNZIPnames
  uzFiles(0 To 99) As String
End Type
'-- Callback Large "String"
Public Type UNZIPCBChar
  ch(32800) As Byte
End Type
'-- Callback Small "String"
Public Type UNZIPCBCh
  ch(256) As Byte
End Type

'-- UNZIP32.DLL DCL Structure
Public Type DCLIST
  ExtractOnlyNewer  As Long    ' 1 = Extract Only Newer/New, Else 0
  SpaceToUnderscore As Long    ' 1 = Convert Space To Underscore, Else 0
  PromptToOverwrite As Long    ' 1 = Prompt To Overwrite Required, Else 0
  fQuiet            As Long    ' 2 = No Messages, 1 = Less, 0 = All
  ncflag            As Long    ' 1 = Write To Stdout, Else 0
  ntflag            As Long    ' 1 = Test Zip File, Else 0
  nvflag            As Long    ' 0 = Extract, 1 = List Zip Contents
  nfflag            As Long    ' 1 = Extract Only Newer Over Existing, Else 0
  nzflag            As Long    ' 1 = Display Zip File Comment, Else 0
  ndflag            As Long    ' 1 = Honor Directories, Else 0
  noflag            As Long    ' 1 = Overwrite Files, Else 0
  naflag            As Long    ' 1 = Convert CR To CRLF, Else 0
  nZIflag           As Long    ' 1 = Zip Info Verbose, Else 0
  C_flag            As Long    ' 1 = Case Insensitivity, 0 = Case Sensitivity
  fPrivilege        As Long    ' 1 = ACL, 2 = Privileges
  Zip               As String  ' The Zip Filename To Extract Files
  ExtractDir        As String  ' The Extraction Directory, NULL If Extracting To Current Dir
End Type

'-- UNZIP32.DLL Userfunctions Structure
Public Type USERFUNCTION
  UZDLLPrnt     As Long     ' Pointer To Apps Print Function
  UZDLLSND      As Long     ' Pointer To Apps Sound Function
  UZDLLREPLACE  As Long     ' Pointer To Apps Replace Function
  UZDLLPASSWORD As Long     ' Pointer To Apps Password Function
  UZDLLMESSAGE  As Long     ' Pointer To Apps Message Function
  UZDLLSERVICE  As Long     ' Pointer To Apps Service Function (Not Coded!)
  TotalSizeComp As Long     ' Total Size Of Zip Archive
  TotalSize     As Long     ' Total Size Of All Files In Archive
  CompFactor    As Long     ' Compression Factor
  NumMembers    As Long     ' Total Number Of All Files In The Archive
  cchComment    As Integer  ' Flag If Archive Has A Comment!
End Type

'-- UNZIP32.DLL Version Structure
Public Type UZPVER
  structlen       As Long         ' Length Of The Structure Being Passed
  flag            As Long         ' Bit 0: is_beta  bit 1: uses_zlib
  beta            As String * 10  ' e.g., "g BETA" or ""
  date            As String * 20  ' e.g., "4 Sep 95" (beta) or "4 September 1995"
  zlib            As String * 10  ' e.g., "1.0.5" or NULL
  unzip(1 To 4)   As Byte         ' Version Type Unzip
  zipinfo(1 To 4) As Byte         ' Version Type Zip Info
  os2dll          As Long         ' Version Type OS2 DLL
  windll(1 To 4)  As Byte         ' Version Type Windows DLL
End Type

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

Enum unxDoingWhat
    nothingtodo = 0
    readcomment = 1
    getfileList = 2
End Enum
Public sUNXMessage As String
Public sUNXPrint As String
Public sUNXFileList As String
Public sUNXComment As String
Private sUNZPWD As String
Private uVbSkip As Integer
Public unxDoingWhatNow As unxDoingWhat



'-- Callback For UNZIP32.DLL - Receive Message Function
Public Sub UZReceiveDLLMessage(ByVal ucsize As Long, _
    ByVal csiz As Long, _
    ByVal cfactor As Integer, _
    ByVal mo As Integer, _
    ByVal dy As Integer, _
    ByVal yr As Integer, _
    ByVal hh As Integer, _
    ByVal mm As Integer, _
    ByVal c As Byte, ByRef fname As UNZIPCBCh, _
    ByRef meth As UNZIPCBCh, ByVal crc As Long, _
    ByVal fCrypt As Byte)

  Dim s0     As String
  Dim xx     As Long

  '-- Always Put This In Callback Routines!
  On Error Resume Next


  s0 = ""

  '-- Do Not Change This For Next!!!
  For xx = 0 To 255
    If fname.ch(xx) = 0 Then Exit For
    s0 = s0 & "%" & Hex$(fname.ch(xx))
  Next
  
  s0 = DecodeUrl(s0, 0)
  
  sUNXMessage = sUNXMessage & s0 & vbCrLf
  
  If unxDoingWhatNow = getfileList Then
    sUNXFileList = sUNXFileList & Chr(0) & s0
  End If

End Sub

'-- Callback For UNZIP32.DLL - Print Message Function

Public Function UZDLLPrnt(ByRef fname As UNZIPCBChar, ByVal x As Long) As Long
    
  Dim s0 As String
  Dim xx As Long

  '-- Always Put This In Callback Routines!
  On Error Resume Next

  s0 = ""

  '-- Gets The UNZIP32.DLL Message For Displaying.
  For xx = 0 To x - 1
    If fname.ch(xx) = 0 Then Exit For
    s0 = s0 & "%" & Hex(fname.ch(xx))
  Next
    s0 = DecodeUrl(s0, 0)

  '-- Assign Zip Information
  If Mid$(s0, 1, 1) = vbLf Then s0 = vbNewLine ' Damn UNIX :-)
  Debug.Print s0
  
  sUNXPrint = sUNXPrint & s0
  
  Select Case unxDoingWhatNow
    Case readcomment
        sUNXComment = s0
        unxDoingWhatNow = 0
     
  End Select


  UZDLLPrnt = 0

End Function

'-- Callback For UNZIP32.DLL - DLL Service Function
Public Function UZDLLServ(ByRef mname As UNZIPCBChar, ByVal x As Long) As Long

    Dim s0 As String
    Dim xx As Long

    '-- Always Put This In Callback Routines!
    On Error Resume Next

    ' Parameter x contains the size of the extracted archive entry.
    ' This information may be used for some kind of progress display...

    s0 = ""
    '-- Get Zip32.DLL Message For processing
    For xx = 0 To UBound(mname.ch)
        If mname.ch(xx) = 0 Then Exit For
        s0 = s0 & Chr$(mname.ch(xx))
    Next
    ' At this point, s0 contains the message passed from the DLL
    ' It is up to the developer to code something useful here :)

    UZDLLServ = 0 ' Setting this to 1 will abort the zip!

End Function

'-- Callback For UNZIP32.DLL - Password Function
Public Function UZDLLPass(ByRef p As UNZIPCBCh, _
  ByVal n As Long, ByRef m As UNZIPCBCh, _
  ByRef name As UNZIPCBCh) As Integer

  Dim prompt     As String
  Dim xx         As Integer
  Dim szpassword As String

  '-- Always Put This In Callback Routines!
  On Error Resume Next

  UZDLLPass = 1

  'If uVbSkip = 1 Then Exit Function

  '-- Get The Zip File Password
  If sUNZPWD = "" Then
    szpassword = InputBox("Please Enter The Password!")
  Else
    szpassword = sUNZPWD
    sUNZPWD = ""
  End If
  

  '-- No Password So Exit The Function
  If Len(szpassword) = 0 Then
    uVbSkip = 1
    Exit Function
  End If

  '-- Zip File Password So Process It
  For xx = 0 To 255
    If m.ch(xx) = 0 Then
      Exit For
    Else
      prompt = prompt & Chr$(m.ch(xx))
    End If
  Next

  For xx = 0 To n - 1
    p.ch(xx) = 0
  Next

  For xx = 0 To Len(szpassword) - 1
    p.ch(xx) = Asc(Mid$(szpassword, xx + 1, 1))
  Next

  p.ch(xx) = 0 ' Put Null Terminator For C

  UZDLLPass = 0

End Function

'-- Callback For UNZIP32.DLL - Report Function To Overwrite Files.
'-- This Function Will Display A MsgBox Asking The User
'-- If They Would Like To Overwrite The Files.
Public Function UZDLLRep(ByRef fname As UNZIPCBChar) As Long

  Dim s0 As String
  Dim xx As Long

  '-- Always Put This In Callback Routines!
  On Error Resume Next

  UZDLLRep = 100 ' 100 = Do Not Overwrite - Keep Asking User
  s0 = ""

  For xx = 0 To 255
    If fname.ch(xx) = 0 Then Exit For
    s0 = s0 & Chr$(fname.ch(xx))
  Next

  '-- This Is The MsgBox Code
  xx = MsgBox("Overwrite " & s0 & "?", vbExclamation & vbYesNoCancel, _
              "VBUnZip32 - File Already Exists!")

  If xx = vbNo Then Exit Function

  If xx = vbCancel Then
    UZDLLRep = 104       ' 104 = Overwrite None
    Exit Function
  End If

  UZDLLRep = 102         ' 102 = Overwrite, 103 = Overwrite All

End Function

Public Function unzErrInfo(iErrCode As unzReturnCode) As String

Select Case iErrCode
    Case 0
    unzErrInfo = "normal; no errors or warnings detected."
    Case 1
    unzErrInfo = "one or more warning errors were encountered, but processing completed " & _
        "successfully anyway.  This includes zipfiles where one or more files " & _
        "was skipped due to unsupported compression method or encryption with an " & _
        "unknown password."
    Case 2
    unzErrInfo = "a generic error in the zipfile format was detected.  Processing may have" & _
        "completed successfully anyway; some broken zipfiles created by other " & _
        "archivers have simple work-arounds."
    Case 3
    unzErrInfo = "a severe error in the zipfile format was detected.  Processing probably" & _
        "failed immediately."
    Case 4
    unzErrInfo = "unzip was unable to allocate memory for one or more buffers during" & _
        "program initialization."
    Case 5
    unzErrInfo = "unzip was unable to allocate memory or unable to obtain a tty to read" & _
        "the decryption password(s)."
    Case 6
    unzErrInfo = "unzip was unable to allocate memory during decompression to disk."
    Case 7
    unzErrInfo = "unzip was unable to allocate memory during in-memory decompression."
    Case 8
    unzErrInfo = "[currently not used]"
    Case 9
    unzErrInfo = "the specified zipfiles were not found."
    Case 10
    unzErrInfo = "invalid options were specified on the command line."
    Case 11
    unzErrInfo = "no matching files were found."
    Case 50
    unzErrInfo = "the disk is (or was) full during extraction."
    Case 51
    unzErrInfo = "the end of the ZIP archive was encountered prematurely."
    Case 80
    unzErrInfo = "the user aborted unzip prematurely with control-C (or similar)"
    Case 81
    unzErrInfo = "testing or extraction of one or more files failed due to unsupported" & _
        "compression methods or unsupported decryption."
    Case 82
    unzErrInfo = "no files were found due to bad decryption password(s).  (If even one file is " & _
        "successfully processed, however, the exit status is 1.)"
End Select
End Function

Public Function infoUnzip(sZipName As String, FileToUnzip As String, UnzipTo As String, bPreserverPath As Boolean, Optional sPWD As String) As unzReturnCode
    Dim unzHandler As New clsUnzip
    With unzHandler
        .sExcludeNames = ""
        .sZipNames = toUnixPath(FileToUnzip)
        .uCaseSensitivity = 0
        .uConvertCR_CRLF = 0
        .uDisplayComment = 0
        .uExtractDir = UnzipTo
        .uExtractList = 0
        .uExtractOnlyNewer = 0
        .uFreshenExisting = 0
        If bPreserverPath Then .uHonorDirectories = 1 Else .uHonorDirectories = 0
        .uOverWriteFiles = 1
        .uPrivilege = 1
        .uPromptOverWrite = 0
        .uQuiet = 2
        .uSpaceUnderScore = 0
        .uTestZip = 0
        .uVerbose = 0
        .uWriteStdOut = 0
        .uZipFileName = toUnixPath(sZipName)
    End With
    unxClearMsg
    sUNZPWD = sPWD
    infoUnzip = unzHandler.UnZip32
End Function
Public Function getCommentText(sZip As String, Optional sPWD As String)

    Dim unzHandler As New clsUnzip
    With unzHandler
        .sExcludeNames = ""
        .sZipNames = ""
        .uCaseSensitivity = 0
        .uConvertCR_CRLF = 0
        .uDisplayComment = 1
        .uExtractDir = Environ("temp")
        .uExtractList = 0
        .uExtractOnlyNewer = 0
        .uFreshenExisting = 0
        .uHonorDirectories = 0
        .uOverWriteFiles = 1
        .uPrivilege = 1
        .uPromptOverWrite = 0
        .uQuiet = 2
        .uSpaceUnderScore = 0
        .uTestZip = 1
        .uVerbose = 0
        .uWriteStdOut = 0
        .uZipFileName = toUnixPath(sZip)
    End With
    sUNZPWD = sPWD
    unxDoingWhatNow = readcomment
    unzHandler.UnZip32
End Function

Public Function getZipFileListText(sZip As String, Optional sPWD As String) As String
    Dim unzHandler As New clsUnzip
    With unzHandler
        .sExcludeNames = ""
        .sZipNames = "*.*"
        .uCaseSensitivity = 1
        .uConvertCR_CRLF = 0
        .uDisplayComment = 0
        .uExtractDir = Environ("temp")
        .uExtractList = 1
        .uExtractOnlyNewer = 0
        .uFreshenExisting = 0
        .uHonorDirectories = 0
        .uOverWriteFiles = 1
        .uPrivilege = 1
        .uPromptOverWrite = 0
        .uQuiet = 2
        .uSpaceUnderScore = 0
        .uTestZip = 0
        .uVerbose = 0
        .uWriteStdOut = 0
        .uZipFileName = toUnixPath(sZip)
    End With
    sUNZPWD = sPWD
    unxDoingWhatNow = getfileList
    unzHandler.UnZip32
    getZipFileListText = sUNXFileList
    unxDoingWhatNow = nothingtodo
End Function

Public Sub unxClearMsg()
    sUNXMessage = ""
    sUNXPrint = ""
    sUNXFileList = ""
    sUNXComment = ""
End Sub
