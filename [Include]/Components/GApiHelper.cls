VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GApiHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1


Public Function CStringToVBString(sCStyle As String) As String
    Dim i As Long
    ' Search for the first null character
    i = InStr(1, sCStyle, vbNullChar)
    If i = 0 Then
        CStringToVBString = sCStyle
    Else
        CStringToVBString = Left$(sCStyle, i - 1)
    End If
End Function

Public Function DWORDPairToDouble(ByRef dwHigh As Long, ByRef dwLow As Long) As Double
    DWORDPairToDouble = (dwHigh * MAXDWORD) + dwLow
End Function

