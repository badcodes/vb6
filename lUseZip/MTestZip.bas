Attribute VB_Name = "TestZipMain"
Option Explicit


Public Sub test()

'
'Dim iz As New cUnzip
'With iz
'.ZipFile = App.Path & "\fortest.zip"
'.AddFileToPreocess "一千零一夜故事集/0/01.htm"
'.UseFolderNames = True
'.UnzipFolder = "c:/"
'End With
''iz.TestZip = True
'iz.unzip

'iz.ZipComment ("ABCDEFG")

Dim uz As New cUnzip
With uz
.ZipFile = App.Path & "\fortest.zip"
.ValidateZipFile .ZipFile
End With



End Sub

