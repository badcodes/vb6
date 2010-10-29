Attribute VB_Name = "MFileData"
Option Explicit

' Required because this is, unknown to all, an ActiveX server
Sub Main()
    Dim test As FTestDictionary
    Set test = New FTestDictionary
    test.Show
End Sub

' No public UDTs allowed in VB5
#If iVBVer > 5 Then
' This function loads all the files matching a given spec into
' a Dictionary.
Function FileDictionary(Optional Spec As String = "*.*", _
                        Optional IncludeDir As Boolean = True _
                        ) As Dictionary

    Dim sName As String, fd As WIN32_FIND_DATA
    Dim hFiles As Long, f As Boolean, dic As Dictionary
    Dim fiTmp As TFileInfo, c As Long
    
    Set dic = New Dictionary
    
    ' Find first file (get handle to find)
    hFiles = FindFirstFile(Spec, fd)
    f = (hFiles <> INVALID_HANDLE_VALUE)
    Do While f
        sName = ByteZToStr(fd.cFileName)
        ' Skip . and .. and unrequested directories
        If (Left$(sName, 1) <> ".") Or _
           ((fd.dwFileAttributes And vbDirectory) And _
           (IncludeDir = False)) Then
            
            ' Create a file info object from file data
            fiTmp.Attribs = fd.dwFileAttributes
            If fd.nFileSizeHigh Then
                fiTmp.Length = -1        ' We can't handle giant files
            Else
                fiTmp.Length = fd.nFileSizeLow
            End If
            fiTmp.LastWrite = Win32ToVBTime(fd.ftLastWriteTime)
            fiTmp.Creation = Win32ToVBTime(fd.ftCreationTime)
            fiTmp.LastAccess = Win32ToVBTime(fd.ftLastAccessTime)
            dic(sName) = fiTmp
        End If
        ' Keep looping until no more files
        f = FindNextFile(hFiles, fd)
    Loop
    f = FindClose(hFiles)
    
    Set FileDictionary = dic
End Function
#End If
