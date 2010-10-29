Attribute VB_Name = "MBookUrl"
Option Explicit
Public Type bookParam
    sParam As String
    sValue As String
End Type
Public Sub test()
Dim a As New CHander
a.wakeUp "book://ssreader/e0?url=http://dl3.lib.tongji.edu.cn/cx210k/30/diskde/de47/25/!00001.pdg&&&&&pages=156&bookname=Â³Ñ¸±ÊÃûÓ¡Æ×"



End Sub
