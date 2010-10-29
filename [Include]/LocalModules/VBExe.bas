Attribute VB_Name = "MVBExe"
Option Explicit
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Public Function IsEXE() As Boolean
    Static bEXE As Boolean
    If Not bEXE Then
        bEXE = True
        Debug.Assert IsEXE() Or True
        IsEXE = bEXE
    End If
    bEXE = False
End Function
