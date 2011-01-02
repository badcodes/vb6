VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MYFAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public count As Integer
Private FCollect() As String
Private Function check(Index As Integer) As Boolean
If Index < 1 Then Exit Function
If Index > count Then Exit Function
check = True
End Function
Public Function name(Index As Integer) As String
If check(Index) = False Then Exit Function
name = FCollect(0, Index)
End Function
Public Function locate(Index As Integer) As String
If check(Index) = False Then Exit Function
locate = FCollect(1, Index)
End Function
Public Sub add(name As String, locate As String)
count = count + 1
If name = "" Then Exit Sub
If locate = "" Then Exit Sub
ReDim Preserve FCollect(1, count) As String
FCollect(0, count) = name
FCollect(1, count) = locate
End Sub
Public Sub del(Index As Integer)
If check(Index) = False Then Exit Sub
For i = Index To count - 1
FCollect(0, i) = FCollect(0, i + 1)
FCollect(1, i) = FCollect(1, i + 1)
Next
count = count - 1
End Sub
Public Sub setfav(Index As Integer, name As String, locate As String)
If check(Index) = False Then Exit Sub
If name = "" Then Exit Sub
If locate = "" Then Exit Sub
FCollect(0, Index) = name
FCollect(1, Index) = locate
End Sub
Public Sub loadfav()
count = Val(GetSetting(App.ProductName, "FavList", "count"))
SaveSetting App.ProductName, "FavList", "count", count
ReDim FCollect(1, count) As String
Dim tempstr As String
Dim pos As Integer
For i = 1 To count
tempstr = GetSetting(App.ProductName, "FavList", Str(i))
pos = InStrRev(tempstr, ",")
If pos > 0 Then
FCollect(0, i) = Left(tempstr, pos - 1)
FCollect(1, i) = Right(tempstr, Len(tempstr) - pos)
Else
i = i - 1
count = count - 1
End If
Next

End Sub
Public Sub savefav()
DeleteSetting App.ProductName, "Favlist"
SaveSetting App.ProductName, "FavList", "count", Str(count)
For i = 1 To count
SaveSetting App.ProductName, "Favlist", Str(i), FCollect(0, i) + "," + FCollect(1, i)
Next

End Sub
