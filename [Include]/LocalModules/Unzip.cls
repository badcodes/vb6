VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUnzip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "clear as the Name below"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'-- Private Variables For Structure Access
Private UZDCL  As DCLIST
Private UZUSER As USERFUNCTION
Private UZVER  As UZPVER

'-- Public Variables For Setting The
'-- UNZIP32.DLL DCLIST Structure
'-- These Must Be Set Before The Actual Call To VBUnZip32
Public uExtractOnlyNewer As Integer  ' 1 = Extract Only Newer/New, Else 0
Public uSpaceUnderScore  As Integer  ' 1 = Convert Space To Underscore, Else 0
Public uPromptOverWrite  As Integer  ' 1 = Prompt To Overwrite Required, Else 0
Public uQuiet            As Integer  ' 2 = No Messages, 1 = Less, 0 = All
Public uWriteStdOut      As Integer  ' 1 = Write To Stdout, Else 0
Public uTestZip          As Integer  ' 1 = Test Zip File, Else 0
Public uExtractList      As Integer  ' 0 = Extract, 1 = List Contents
Public uFreshenExisting  As Integer  ' 1 = Update Existing by Newer, Else 0
Public uDisplayComment   As Integer  ' 1 = Display Zip File Comment, Else 0
Public uHonorDirectories As Integer  ' 1 = Honor Directories, Else 0
Public uOverWriteFiles   As Integer  ' 1 = Overwrite Files, Else 0
Public uConvertCR_CRLF   As Integer  ' 1 = Convert CR To CRLF, Else 0
Public uVerbose          As Integer  ' 1 = Zip Info Verbose
Public uCaseSensitivity  As Integer  ' 1 = Case Insensitivity, 0 = Case Sensitivity
Public uPrivilege        As Integer  ' 1 = ACL, 2 = Privileges, Else 0
Public uZipFileName      As String   ' The Zip File Name
Public uExtractDir       As String   ' Extraction Directory, Null If Current Directory

'-- Public Program Variables
Public uZipNumber    As Long         ' Zip File Number
Public uNumberFiles  As Long         ' Number Of Files
Public uNumberXFiles As Long         ' Number Of Extracted Files
Public uZipMessage   As String       ' For Zip Message
Public uZipInfo      As String       ' For Zip Information
Private uZipNames     As UNZIPnames   ' Names Of Files To Unzip
Private uExcludeNames As UNZIPnames   ' Names Of Zip Files To Exclude
Public sZipNames As String
Public sExcludeNames As String
Public uVbSkip       As Integer      ' For DLL Password Function

'-- Puts A Function Pointer In A Structure
'-- For Callbacks.
Public Function FnPtr(ByVal lp As Long) As Long

  FnPtr = lp

End Function


'-- ASCIIZ To String Function
Private Function szTrim(szString As String) As String

  Dim pos As Long

  pos = InStr(szString, vbNullChar)

  Select Case pos
    Case Is > 1
      szTrim = Trim$(Left$(szString, pos - 1))
    Case 1
      szTrim = ""
    Case Else
      szTrim = Trim$(szString)
  End Select

End Function

'-- Main UNZIP32.DLL UnZip32 Subroutine
'-- (WARNING!) Do Not Change!
Public Function UnZip32() As unzReturnCode

  Dim retcode As Long

  '-- Set The UNZIP32.DLL Options
  '-- (WARNING!) Do Not Change
  UZDCL.ExtractOnlyNewer = uExtractOnlyNewer ' 1 = Extract Only Newer/New
  UZDCL.SpaceToUnderscore = uSpaceUnderScore ' 1 = Convert Space To Underscore
  UZDCL.PromptToOverwrite = uPromptOverWrite ' 1 = Prompt To Overwrite Required
  UZDCL.fQuiet = uQuiet                      ' 2 = No Messages 1 = Less 0 = All
  UZDCL.ncflag = uWriteStdOut                ' 1 = Write To Stdout
  UZDCL.ntflag = uTestZip                    ' 1 = Test Zip File
  UZDCL.nvflag = uExtractList                ' 0 = Extract 1 = List Contents
  UZDCL.nfflag = uFreshenExisting            ' 1 = Update Existing by Newer
  UZDCL.nzflag = uDisplayComment             ' 1 = Display Zip File Comment
  UZDCL.ndflag = uHonorDirectories           ' 1 = Honour Directories
  UZDCL.noflag = uOverWriteFiles             ' 1 = Overwrite Files
  UZDCL.naflag = uConvertCR_CRLF             ' 1 = Convert CR To CRLF
  UZDCL.nZIflag = uVerbose                   ' 1 = Zip Info Verbose
  UZDCL.C_flag = uCaseSensitivity            ' 1 = Case insensitivity, 0 = Case Sensitivity
  UZDCL.fPrivilege = uPrivilege              ' 1 = ACL 2 = Priv
  UZDCL.Zip = uZipFileName                   ' ZIP Filename
  UZDCL.ExtractDir = uExtractDir             ' Extraction Directory, NULL If Extracting
                                             ' To Current Directory

  '-- Set Callback Addresses
  '-- (WARNING!!!) Do Not Change
  UZUSER.UZDLLPrnt = FnPtr(AddressOf modUnzip.UZDLLPrnt)
  UZUSER.UZDLLSND = 0&    '-- Not Supported
  UZUSER.UZDLLREPLACE = FnPtr(AddressOf modUnzip.UZDLLRep)
  UZUSER.UZDLLPASSWORD = FnPtr(AddressOf modUnzip.UZDLLPass)
  UZUSER.UZDLLMESSAGE = FnPtr(AddressOf modUnzip.UZReceiveDLLMessage)
  UZUSER.UZDLLSERVICE = FnPtr(AddressOf modUnzip.UZDLLServ)

  '-- Set UNZIP32.DLL Version Space
  '-- (WARNING!!!) Do Not Change
  With UZVER
    .structlen = Len(UZVER)
    .beta = space$(9) & vbNullChar
    .date = space$(19) & vbNullChar
    .zlib = space$(9) & vbNullChar
  End With
  
  Dim sArrTmp() As String
  Dim lcount As Long
  Dim l As Long
  With uZipNames
    .uzFiles(0) = ""
    sArrTmp = Split(sZipNames, Chr(0))
    lcount = UBound(sArrTmp)
    If lcount > 98 Then lcount = 98
    For l = 0 To lcount
    .uzFiles(l) = sArrTmp(l)
    Next
    uNumberFiles = lcount + 1
    If lcount >= 0 Then .uzFiles(lcount + 1) = vbNullString
  End With
  
  With uExcludeNames
    .uzFiles(0) = ""
    sArrTmp = Split(sExcludeNames, Chr(0))
    lcount = UBound(sArrTmp)
    If lcount > 98 Then lcount = 98
    For l = 0 To lcount
    .uzFiles(l) = sArrTmp(l)
    Next
    uNumberXFiles = lcount + 1
    If lcount >= 0 Then .uzFiles(lcount + 1) = vbNullString
  End With

  '-- Get Version
  Call UzpVersion2(UZVER)
  
  '--------------------------------------
  '-- You Can Change This For Displaying
  '-- The Version Information!
  '--------------------------------------
'  MsgStr$ = "DLL Date: " & szTrim(UZVER.date)
'  MsgStr$ = MsgStr$ & vbNewLine$ & "Zip Info: " & Hex$(UZVER.zipinfo(1)) & "." & _
'       Hex$(UZVER.zipinfo(2)) & Hex$(UZVER.zipinfo(3))
'  MsgStr$ = MsgStr$ & vbNewLine$ & "DLL Version: " & Hex$(UZVER.windll(1)) & "." & _
'       Hex$(UZVER.windll(2)) & Hex$(UZVER.windll(3))
'  MsgStr$ = MsgStr$ & vbNewLine$ & "--------------"
  '-- End Of Version Information.

  '-- Go UnZip The Files! (Do Not Change Below!!!)
  '-- This Is The Actual UnZip Routine
  retcode = Wiz_SingleEntryUnzip(uNumberFiles, uZipNames, uNumberXFiles, _
                                 uExcludeNames, UZDCL, UZUSER)
  '---------------------------------------------------------------

  '-- If There Is An Error Display A MsgBox!
  If retcode <> 0 Then Debug.Print unzErrInfo(retcode)
  UnZip32 = retcode
  

    '-- You Can Change This As Needed!
    '-- For Compression Information
    '  MsgStr$ = MsgStr$ & vbNewLine & "Only Shows If uExtractList = 1 List Contents"
    '  MsgStr$ = MsgStr$ & vbNewLine & "--------------"
    '  MsgStr$ = MsgStr$ & vbNewLine & "Comment         : " & UZUSER.cchComment
    '  MsgStr$ = MsgStr$ & vbNewLine & "Total Size Comp : " & UZUSER.TotalSizeComp
    '  MsgStr$ = MsgStr$ & vbNewLine & "Total Size      : " & UZUSER.TotalSize
    '  MsgStr$ = MsgStr$ & vbNewLine & "Compress Factor : %" & UZUSER.CompFactor
    '  MsgStr$ = MsgStr$ & vbNewLine & "Num Of Members  : " & UZUSER.NumMembers
    '  MsgStr$ = MsgStr$ & vbNewLine & "--------------"

End Function
