VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZhBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Need modIni(ini.bas)

Private iBMpartCount As Integer
Private iBMcount As Integer
Private sArrBookmark() As String

Function newBookmark(iBookmarkPartAmount As Integer) As Boolean
Me.selfReset
iBMpartCount = iBookmarkPartAmount
End Function

Function Count() As Integer
Count = iBMcount
End Function

Function PartCount() As Integer
PartCount = iBMpartCount
End Function

Sub Add(sArrSource() As String)

Dim iArrAmount As Integer
Dim i As Integer
iArrAmount = UBound(sArrSource) + 1
If iArrAmount = 0 Then Exit Sub
If iArrAmount > iBMpartCount Then iArrAmount = iBMpartCount

ReDim Preserve sArrBookmark(iBMcount) As String

For i = 0 To iArrAmount - 1
sArrBookmark(iBMcount, i) = sArrSource(i)
Next

iBMcount = iBMcount + 1

End Sub
Sub Remove(iBookmarkIndex As Integer)
Dim i As Integer
If iBookmarkIndex < 0 Then Exit Sub
If iBookmarkIndex >= iBMcount Then Exit Sub
For i = iBookmarkIndex To iBMcount - 2
sArrBookmark(i) = sArrBookmark(i + 1)
Next
End Sub

Sub SetAs(iBookmarkIndex As Integer, sArrSource() As String)

Dim iArrAmount As Integer
Dim i As Integer

If iBookmarkIndex < 0 Then Exit Sub
If iBookmarkIndex >= iBMcount Then Exit Sub
iArrAmount = UBound(sArrSource) + 1
If iArrAmount = 0 Then Exit Sub
If iArrAmount > iBMpartCount Then iArrAmount = iBMpartCount

For i = 0 To iBMpartCount - 1
sArrBookmark(iBookmarkIndex, i) = ""
Next

For i = 0 To iArrAmount - 1
sArrBookmark(iBookmarkIndex, i) = sArrSource(i)
Next

End Sub
Function getBookmark(iBMindex As Integer, iBMPartIndex As Integer) As String

If iBMindex > iBMcount - 1 Then Exit Function
If iBMPartIndex > iBMpartCount - 1 Then Exit Function
getBookmark = sArrBookmark(iBMindex, iBMPartIndex)

End Function

Function saveToFile(sBMfile As String)
Dim i As Integer
Dim j As Integer


iniSaveSetting sBMfile, "Head", "BookmarkCount", iBMpartCount
iniSaveSetting sBMfile, "Head", "BMPartCount", iBMcount

For i = 0 To iBMcount - 1
    For j = 0 To iBMpartCount - 1
    iniSaveSetting sBMfile, "Body", getBookmarkKeyName(i, j), sArrBookmark(i, j)
    Next
Next

End Function

Function loadFromFile(sBMfile As String)
Dim i As Integer
Dim j As Integer
Dim sBookmarkKeyName As String

iBMcount = Val(iniGetSetting(sBMfile, "Head", "BookmarkCouont"))
iBMpartCount = Val(iniGetSetting(sBMfile, "Head", "BMPartCount"))
ReDim sArrBookmark(iBMcount - 1, iBMpartCount - 1) As String

For i = 0 To iBMcount - 1
    For j = 0 To iBMpartCount - 1
    sArrBookmark(i, j) = iniGetSetting(sBMfile, "Body", getBookmarkKeyName(i, j))
    Next
Next

Me.bBMloaded = True

End Function

Sub selfReset()

iBMpartCount = 0
iBMcount = 0
bBMloaded = False
Erase sArrBookmark

End Sub

Private Function getBookmarkKeyName(iBMindex As Integer, iBMPartIndex As Integer) As String
    getBookmarkKeyName = "BM_" & Str(iBMindex) & "_" & Str(iBMPartIndex)
End Function
