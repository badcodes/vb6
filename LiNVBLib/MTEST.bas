Attribute VB_Name = "MTEST"
Option Explicit
Public Sub test()
Dim gcf As New CRegistry
gcf.ClassKey = HKEY_CLASSES_ROOT
gcf.SectionKey = ".zhtm"
'Debug.Print .ToString
End Sub
