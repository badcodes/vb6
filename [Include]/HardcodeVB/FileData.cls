VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' No UDTs allowed in VB5
#If iVBVEr > 5 Then
' UDT for Variant must be in a public class of a server,
' whether you want it there or not
Type TFileInfo
    Attribs As Long
    LastWrite As Date
    Creation As Date
    LastAccess As Date
    Length As Long
End Type

' Can't do this
'Public Info As TFileInfo

Private fi As TFileInfo

Property Get FileInfo() As TFileInfo
    ' Validate output here
    FileInfo = fi
End Property

Property Let FileInfo(fiA As TFileInfo)
    ' Validate input here
    fi = fiA
End Property
#End If
