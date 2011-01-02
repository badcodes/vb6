Attribute VB_Name = "MInfoZipShared"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Enum LOCALEID
    ZH_CN = &H804
    ZH_TW = &H404
    ZH_Hans = &H4
    ZH_Hant = &H7C04
    EN = &H9
    EN_US = &H409
    JA_JP = &H411
    JA = &H11
End Enum

Public Function strConvX(srcString As Variant, Conversion As VbStrConv, Optional LCID As LOCALEID) As String
    strConvX = StrConv(srcString, Conversion, LCID)
End Function

Public Function CBytesToStr(ByRef CBytes() As Byte) As String

    Dim lUB As Long, lLb As Long
    Dim iPos As Long
    Dim bTemp() As Byte
    Dim l As Long
    
    lUB = UBound(CBytes)
    lLb = LBound(CBytes)
    
    For iPos = lLb To lUB
        If CBytes(iPos) = 0 Then Exit For
    Next
    
    If iPos = 0 Then
        CBytesToStr = strConvX(CBytes, vbUnicode, ZH_CN)
    ElseIf iPos = lLb Then
        CBytesToStr = ""
    Else
        ReDim bTemp(lLb To iPos - 1)
        CopyMemory bTemp(lLb), CBytes(lLb), iPos - lLb
        CBytesToStr = strConvX(bTemp, vbUnicode, ZH_CN)
    End If
    Debug.Print CBytesToStr
    
End Function

Public Sub StrToCBytes(ByVal strUnicode As String, ByRef CBytes() As Byte)

    Dim lUB As Long, lLb As Long
    Dim bTemp() As Byte
    Dim lSize As Long
    
    lUB = UBound(CBytes)
    lLb = LBound(CBytes)
    
    bTemp = strConvX(strUnicode, vbFromUnicode, ZH_CN)
    
    lSize = UBound(bTemp) + 1
    ReDim Preserve bTemp(lSize)
    bTemp(lSize) = 0
    
    If lSize > lUB - lLb Then
        lSize = lUB - lLb
        bTemp(lSize) = 0
    End If
    
    CopyMemory CBytes(lLb), bTemp(0), lSize + 1
       
End Sub

Public Function CleanZipFilename(sInCome) As String

Dim sFilenameClean As String
Dim iPos As Long, iStart As Long, iEnd As Long
Dim charNow As String
iEnd = Len(sInCome)
iStart = 1
For iPos = iStart To iEnd
    charNow = Mid$(sInCome, iPos, 1)
    Select Case charNow
        Case "\"
            sFilenameClean = sFilenameClean & "/"
        Case "["
            sFilenameClean = sFilenameClean & "[[]"
       ' Case "]"
           ' sFilenameClean = sFilenameClean & "[]]"
        Case Else
            sFilenameClean = sFilenameClean & charNow
    End Select
 Next

CleanZipFilename = sFilenameClean

End Function

'Public Function CBytesToStr(ByRef CBytes() As Byte) As String
'    CBytesToStr = mShareFunction.CBytesToStr(CBytes())
'End Function
'Public Sub StrToCBytes(ByVal strUnicode As String, ByRef CBytes() As Byte)
'    Call mShareFunction.StrToCBytes(strUnicode, CBytes())
'End Sub

