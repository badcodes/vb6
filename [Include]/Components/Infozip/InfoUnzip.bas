Attribute VB_Name = "mUnzip"
Option Explicit

' ======================================================================================
' Name:     mUnzip
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     1 December 2000
'
' Requires: Info-ZIP's Unzip32.DLL v5.40, renamed to vbuzip10.dll
'           cUnzip.cls
'
' Copyright © 2000 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Part of the implementation of cUnzip.cls, a class which gives a
' simple interface to Info-ZIP's excellent, free unzipping library
' (Unzip32.DLL).
'
' This sample uses decompression code by the Info-ZIP group.  The
' original Info-Zip sources are freely available from their website
' at
'     http://www.cdrcom.com/pubs/infozip/
'
' Please ensure you visit the site and read their free source licensing
' information and requirements before using their code in your own
' application.
'
' ======================================================================================

' argv
Private Type UNZIPnames
    S(0 To 1023) As String
End Type

' Callback large "string" (sic)
Private Type CBChar
    ch(0 To 32800) As Byte
End Type

' Callback small "string" (sic)
Private Type CBCh
    ch(0 To 255) As Byte
End Type

' DCL structure
Public Type DCLIST
   ExtractOnlyNewer As Long      ' 1 to extract only newer
   SpaceToUnderscore As Long     ' 1 to convert spaces to underscore
   PromptToOverwrite As Long     ' 1 if overwriting prompts required
   fQuiet As Long                ' 0 = all messages, 1 = few messages, 2 = no messages
   ncflag As Long                ' write to stdout if 1
   ntflag As Long                ' test zip file
   nvflag As Long                ' verbose listing
   nUflag As Long                ' "update" (extract only newer/new files)
   nzflag As Long                ' display zip file comment
   ndflag As Long                ' all args are files/dir to be extracted
   noflag As Long                ' 1 if always overwrite files
   naflag As Long                ' 1 to do end-of-line translation
   nZIflag As Long               ' 1 to get zip info
   C_flag As Long                ' 1 to be case insensitive
   fPrivilege As Long            ' zip file name
   lpszZipFN As String           ' directory to extract to.
   lpszExtractDir As String
End Type

Private Type USERFUNCTION
   ' Callbacks:
   lptrPrnt As Long           ' Pointer to application's print routine
   lptrSound As Long          ' Pointer to application's sound routine.  NULL if app doesn't use sound
   lptrReplace As Long        ' Pointer to application's replace routine.
   lptrPassword As Long       ' Pointer to application's password routine.
   lptrMessage As Long        ' Pointer to application's routine for
                              ' displaying information about specific files in the archive
                              ' used for listing the contents of the archive.
   lptrService As Long        ' callback function designed to be used for allowing the
                              ' app to process Windows messages, or cancelling the operation
                              ' as well as giving option of progress.  If this function returns
                              ' non-zero, it will terminate what it is doing.  It provides the app
                              ' with the name of the archive member it has just processed, as well
                              ' as the original size.
                              
   ' Values filled in after processing:
   lTotalSizeComp As Long     ' Value to be filled in for the compressed total size, excluding
                              ' the archive header and central directory list.
   lTotalSize As Long         ' Total size of all files in the archive
   lCompFactor As Long        ' Overall archive compression factor
   lNumMembers As Long        ' Total number of files in the archive
   cchComment As Integer      ' Flag indicating whether comment in archive.
End Type

Public Type ZIPVERSIONTYPE
   major As Byte
   minor As Byte
   patchlevel As Byte
   not_used As Byte
End Type

Public Type UZPVER
    structlen As Long         ' Length of structure
    flag As Long              ' 0 is beta, 1 uses zlib
    betalevel As String * 10  ' e.g "g BETA"
    date As String * 20       ' e.g. "4 Sep 95" (beta) or "4 September 1995"
    zlib As String * 10       ' e.g. "1.0.5 or NULL"
    unzip As ZIPVERSIONTYPE
    zipinfo As ZIPVERSIONTYPE
    os2dll As ZIPVERSIONTYPE
    windll As ZIPVERSIONTYPE
End Type

Private Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" _
  (ByVal ifnc As Long, ByRef ifnv As UNZIPnames, _
   ByVal xfnc As Long, ByRef xfnv As UNZIPnames, _
   dcll As DCLIST, Userf As USERFUNCTION) As Long
   
'Public Declare Sub UzpVersion2 Lib "unzip32.dll" (uzpv As UZPVER)
Public Declare Function Wiz_Validate Lib "unzip32.dll" (ByVal sArchive As String, ByVal AllCodes As Long) As Long
' Object for callbacks:
Private m_cUnzip As cUnzip
Private m_bCancel As Boolean
Private m_bBusy As Boolean

Private Function plAddressOf(ByVal lPtr As Long) As Long
   ' VB Bug workaround fn
   plAddressOf = lPtr
End Function

Private Sub UnzipMessageCallBack( _
      ByVal ucsize As Long, _
      ByVal csiz As Long, _
      ByVal cfactor As Integer, _
      ByVal mo As Integer, _
      ByVal dy As Integer, _
      ByVal yr As Integer, _
      ByVal hh As Integer, _
      ByVal mm As Integer, _
      ByVal c As Byte, _
      ByRef fname As CBCh, _
      ByRef meth As CBCh, _
      ByVal Crc As Long, _
      ByVal fCrypt As Byte _
   )
Dim sFilename As String
'Dim sFolder As String
Dim dDate As Date
Dim sMethod As String
Dim iPos As Long

   On Error Resume Next
   'MsgBox "MessageCallBack Begin"
        
     sFilename = CBytesToStr(fname.ch)
     sMethod = CBytesToStr(meth.ch)
     dDate = DateSerial(yr, mo, hh)
     dDate = dDate + TimeSerial(hh, mm, 0)
      
   'Debug.Print "MessageCallBack End"
    If m_cUnzip.isGetZipItems Then
        m_cUnzip.AddZipItem sFilename, dDate, ucsize, csiz, Crc, ((fCrypt And 64) = 64), cfactor, sMethod
    End If
      'MsgBox "Call .AddFile Done"

   
End Sub

Private Function UnzipPrintCallback( _
      ByRef fname As CBChar, _
      ByVal x As Long _
   ) As Long
Dim iPos As Long
Dim sMsg As String
   On Error Resume Next
   
   ' Check we've got a message:
   If x > 1 And x < 32000 Then

    sMsg = CBytesToStr(fname.ch)
    
    'W_lstrcpynPtrStr sMsg, fname.ch(0), 10
      If m_cUnzip.isGetComment = True Then
            m_cUnzip.isGetComment = False
            m_cUnzip.Comment = sMsg
      End If
      
      Debug.Print "UnzipPrintCallback:" & sMsg
      m_cUnzip.ProgressReport sMsg
      
   End If
   UnzipPrintCallback = 0
End Function

Private Function UnzipPasswordCallBack( _
      ByRef pwd As CBCh, _
      ByVal x As Long, _
      ByRef s2 As CBCh, _
      ByRef name As CBCh _
   ) As Long

Dim bCancel As Boolean
Dim sPassword As String
Dim b() As Byte
Dim lSize As Long
Dim sName As String
Dim iPos As Long

On Error Resume Next

   ' The default:
   UnzipPasswordCallBack = 1
    
   If m_bCancel Then
      Exit Function
   End If
   
   ' Ask for password:
   sName = CBytesToStr(name.ch)
   m_cUnzip.PasswordRequest sPassword, sName, bCancel
      
   sPassword = Trim$(sPassword)
   
   ' Cancel out if no useful password:
   If bCancel Or Len(sPassword) = 0 Then
      m_bCancel = True
      Exit Function
   End If
   
   MInfoZipShared.StrToCBytes sPassword, pwd.ch
   
'   ' Put password into return parameter:
'   lSize = Len(sPassword)
'   If lSize > 254 Then
'      lSize = 254
'   End If
'
'   b = StrConv(sPassword, vbFromUnicode)
'   CopyMemory pwd.ch(0), b(0), lSize
'   pwd.ch(lSize) = 0
   
   ' Ask UnZip to process it:
   UnzipPasswordCallBack = 0
       
End Function

Private Function UnzipReplaceCallback(ByRef fname As CBChar) As Long
Dim eResponse As EUZOverWriteResponse
Dim iPos As Long
Dim sFile As String

   On Error Resume Next
   eResponse = euzDoNotOverwrite
   
   ' Extract the filename:
   sFile = MInfoZipShared.CBytesToStr(fname.ch)
   
'   sFile = StrConv(fname.ch, vbUnicode)
'   iPos = InStr(sFile, vbNullChar)
'   If (iPos > 1) Then
'      sFile = Left$(sFile, iPos - 1)
'   End If
   
   ' No backslashes:
   'sFile = Replace$(sFile, "/", "\")
   
   ' Request the overwrite request:
   m_cUnzip.OverwriteRequest sFile, eResponse
   
   ' Return it to the zipping lib
   UnzipReplaceCallback = eResponse
   
End Function

Private Function UnZipServiceCallback(ByRef mname As CBChar, ByVal x As Long) As Long

Dim sInfo As String
Dim bCancel As Boolean
    
'-- Always Put This In Callback Routines!
On Error Resume Next
    
   ' Check we've got a message:
   If x > 1 And x < 32000 Then
   
      sInfo = MInfoZipShared.CBytesToStr(mname.ch)
      m_cUnzip.Service sInfo, bCancel
      
      If bCancel Then
         UnZipServiceCallback = 1
      Else
         UnZipServiceCallback = 0
      End If
   End If
   
End Function

Public Function VBUnzip( _
      cUnzipObject As cUnzip, _
      tDCL As DCLIST, _
      iIncCount As Long, _
      sInc() As String, _
      iExCount As Long, _
      sExc() As String _
   ) As Long
Dim tUser As USERFUNCTION
Dim lR As Long
Dim tInc As UNZIPnames
Dim tExc As UNZIPnames
Dim i As Long

Do Until m_bBusy = False
DoEvents
Loop

m_bBusy = True
On Error GoTo ErrorHandler

   Set m_cUnzip = cUnzipObject
   ' Set Callback addresses
   tUser.lptrPrnt = plAddressOf(AddressOf UnzipPrintCallback)
   tUser.lptrSound = 0& ' not supported
   tUser.lptrReplace = plAddressOf(AddressOf UnzipReplaceCallback)
   tUser.lptrPassword = plAddressOf(AddressOf UnzipPasswordCallBack)
   tUser.lptrMessage = plAddressOf(AddressOf UnzipMessageCallBack)
   tUser.lptrService = plAddressOf(AddressOf UnZipServiceCallback)
        
   ' Set files to include/exclude:
   Dim lStart As Long
   Dim lEnd As Long
   
   If (iIncCount > 0) Then
      lStart = LBound(sInc())
      lEnd = iIncCount - 1 + lStart
      For i = lStart To lEnd
         tInc.S(i - lStart) = sInc(i)
      Next i
      tInc.S(lEnd - lStart + 1) = vbNullChar
   Else
      tInc.S(0) = vbNullChar
   End If
   
   If (iExCount > 0) Then
      lStart = LBound(sExc())
      lEnd = iExCount - 1 + lStart
      For i = lStart To lEnd
         tExc.S(i - lStart) = sExc(i)
      Next i
      tExc.S(lEnd - lStart + 1) = vbNullChar
   Else
      tExc.S(0) = vbNullChar
   End If
   m_bCancel = False
   VBUnzip = Wiz_SingleEntryUnzip(iIncCount, tInc, iExCount, tExc, tDCL, tUser)
   m_bBusy = False
    'Debug.Print "--------------"
    'Debug.Print MYUSER.cchComment
    'Debug.Print MYUSER.TotalSizeComp
    'Debug.Print MYUSER.TotalSize
    'Debug.Print MYUSER.CompFactor
    'Debug.Print MYUSER.NumMembers
    'Debug.Print "--------------"
   Set m_cUnzip = Nothing
   Exit Function
   
ErrorHandler:
   m_bBusy = False
Dim lErr As Long, sErr As String
   lErr = Err.Number: sErr = Err.Description
   VBUnzip = -1
   Set m_cUnzip = Nothing
   Err.Raise lErr, App.EXEName & ".VBUnzip", sErr
   Exit Function

End Function

'Public Function unzErrInfo(iErrCode As unzReturnCode) As String
'
'Select Case iErrCode
'    Case 0
'    unzErrInfo = "normal; no errors or warnings detected."
'    Case 1
'    unzErrInfo = "one or more warning errors were encountered, but processing completed " & _
'        "successfully anyway.  This includes zipfiles where one or more files " & _
'        "was skipped due to unsupported compression method or encryption with an " & _
'        "unknown password."
'    Case 2
'    unzErrInfo = "a generic error in the zipfile format was detected.  Processing may have" & _
'        "completed successfully anyway; some broken zipfiles created by other " & _
'        "archivers have simple work-arounds."
'    Case 3
'    unzErrInfo = "a severe error in the zipfile format was detected.  Processing probably" & _
'        "failed immediately."
'    Case 4
'    unzErrInfo = "unzip was unable to allocate memory for one or more buffers during" & _
'        "program initialization."
'    Case 5
'    unzErrInfo = "unzip was unable to allocate memory or unable to obtain a tty to read" & _
'        "the decryption password(s)."
'    Case 6
'    unzErrInfo = "unzip was unable to allocate memory during decompression to disk."
'    Case 7
'    unzErrInfo = "unzip was unable to allocate memory during in-memory decompression."
'    Case 8
'    unzErrInfo = "[currently not used]"
'    Case 9
'    unzErrInfo = "the specified zipfiles were not found."
'    Case 10
'    unzErrInfo = "invalid options were specified on the command line."
'    Case 11
'    unzErrInfo = "no matching files were found."
'    Case 50
'    unzErrInfo = "the disk is (or was) full during extraction."
'    Case 51
'    unzErrInfo = "the end of the ZIP archive was encountered prematurely."
'    Case 80
'    unzErrInfo = "the user aborted unzip prematurely with control-C (or similar)"
'    Case 81
'    unzErrInfo = "testing or extraction of one or more files failed due to unsupported" & _
'        "compression methods or unsupported decryption."
'    Case 82
'    unzErrInfo = "no files were found due to bad decryption password(s).  (If even one file is " & _
'        "successfully processed, however, the exit status is 1.)"
'End Select
'End Function

