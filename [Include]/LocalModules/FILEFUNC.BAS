Attribute VB_Name = "MFILEFUNC"
Option Explicit

Function FileExists(ByVal FileName As String) As Boolean
Dim Temp$

    'Set Default
    FileExists = True
    
    'Set up error handler
On Error Resume Next

    'Attempt to grab date and time
    Temp$ = FileDateTime(FileName)

    'Process errors
    Select Case Err
        Case 53, 76, 68   'File Does Not Exist
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, vbOKOnly, "Error"
                End
            End If
    End Select
End Function

Function AddPathToFile(ByVal sPathIn As String, ByVal sFileNameIn As String) As String
'*******************************************************************
'
'  PURPOSE: Takes a path (including Drive letter and any subdirs) and
'           concatenates the file name to path. Path may be empty, path
'           may or may not have an ending backslash '\'.  No validation
'           or existance is check on path or file.
'
'  INPUTS:  sPathIn - Path to use
'           sFileNameIn - Filename to use
'
'
'  OUTPUTS:  N/A
'
'  RETURNS:  Path concatenated to File.
'
'*******************************************************************
Dim sPath As String
Dim sFileName As String


    'Remove any leading or trailing spaces
    sPath = Trim$(sPathIn)
    sFileName = Trim$(sFileNameIn)

    If sPath = "" Then
       AddPathToFile = sFileName
    Else
       If Right$(sPath, 1) = "\" Then
         AddPathToFile = sPath & sFileName
       Else
         AddPathToFile = sPath & "\" & sFileName
       End If
    End If

End Function
Function ExtractFileName(sFileName As Variant) As String
'*******************************************************************
'
'  PURPOSE: This returns just a file name from a full/partial path.
'
'  INPUTS:  sFileName - String Data to remove path from.
'
'  OUTPUTS: N/A
'
'  RETURNS: This function returns all the characters from right to the
'           first \.  Does NOT check validity of the filename....
'
'*******************************************************************
Dim nIdx As Integer


    For nIdx = Len(sFileName) To 1 Step -1
        If Mid$(sFileName, nIdx, 1) = "\" Then
            ExtractFileName = Mid$(sFileName, nIdx + 1)
            Exit Function
        End If
    Next nIdx

    ExtractFileName = sFileName

End Function

Function ExtractPath(sFileName) As String
'*******************************************************************
'
'  PURPOSE: This returns just a path name from a full/partial path.
'
'  INPUTS:  sFileName - String Data to remove file from.
'
'  OUTPUTS: N/A
'
'  RETURNS: This function returns all the characters from left to the last
'           first \.  Does NOT check validity of the filename/Path....
'*******************************************************************
Dim nIdx As Integer


    For nIdx = Len(sFileName) To 1 Step -1
       If Mid$(sFileName, nIdx, 1) = "\" Then
          ExtractPath = Mid$(sFileName, 1, nIdx)
          Exit Function
       End If
    Next nIdx
    
    ExtractPath = sFileName

End Function


