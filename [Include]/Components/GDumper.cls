VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GDumper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Dumper_HexTable(ByRef Data() As Byte, Optional vStart As Long = -1, Optional vLength As Long = -1) As String
    On Error Resume Next
    Dim Result As String
    If vStart < 1 Then vStart = 0
    If vLength < 1 Then vLength = UBound(Data())
    Dim vEnd As Long
    vEnd = vStart + vLength - 1
    
    Dim c As Long
    c = 0
    Do While c < vLength
        Dim i As Long
        Dim iLine As Long
        Dim iStart As Long
        Dim iEnd As Long
        iStart = iLine * 16
        iEnd = iStart + 16 - 1
        If iEnd > vEnd Then iEnd = vEnd
        For i = iStart To iEnd
            Debug.Print Hex(Data(i)); " ";
        Next
        For i = iEnd To (16 - (iEnd - iStart))
            Debug.Print "   ";
        Next
        Debug.Print "  ";
        For i = iStart To iEnd
            Debug.Print Chr$(Data(i)); " ";
        Next
        Debug.Print ""
        c = c + iEnd - iStart + 1
        iLine = iLine + 1
    Loop
    
End Function

Public Sub DumpArray(vArray As Variant, Optional level As Long = 1)
    Dim i As Long
    Dim pL As Long
    Dim pU As Long
    If level <= 1 Then
        level = 1
        pL = ArrayBound(vArray)
        pU = ArrayBound(vArray, True)
        For i = pL To pU
            Debug.Print "VArray("; String(level, " , "); i; ")="; vArray(i)
        Next
    ElseIf level = 2 Then
        DumpArray2 vArray
    ElseIf level = 3 Then
        DumpArray3 vArray
    Else
        MsgBox "Unsupport Level", vbCritical
    End If
End Sub

'CSEH: ErrDebugPrint
Private Sub DumpArray3(vArray As Variant)
    '<EhHeader>
    On Error GoTo DumpArray3_Err
    '</EhHeader>
        Dim i As Long
        Dim iL As Long
        Dim iU As Long
        Dim j As Long
        Dim jL As Long
        Dim jU As Long
        Dim k As Long
        Dim kL As Long
        Dim kU As Long
        iL = LBound(vArray, 1)
        iU = UBound(vArray, 1)
        jL = LBound(vArray, 2)
        jU = UBound(vArray, 2)
        kL = LBound(vArray, 3)
        kU = UBound(vArray, 3)
        For i = iL To iU
            For j = jL To jU
                For k = kL To kU
                    Debug.Print "Array("; i; ","; j; ","; k; ") = "; vArray(i, j, k)
                Next
            Next
        Next

    '<EhFooter>
    Exit Sub
DumpArray3_Err:
    Debug.Print "GetSSLib.MDumper.DumpArray3:Error " & Err.Description
    
    '</EhFooter>
End Sub
'CSEH: ErrDebugPrint
Private Sub DumpArray2(vArray As Variant)
    '<EhHeader>
    On Error GoTo DumpArray2_Err
    '</EhHeader>
        Dim i As Long
        Dim iL As Long
        Dim iU As Long
        Dim j As Long
        Dim jL As Long
        Dim jU As Long
        iL = LBound(vArray, 1)
        iU = UBound(vArray, 1)
        jL = LBound(vArray, 2)
        jU = UBound(vArray, 2)
        For i = iL To iU
            For j = jL To jU
                Debug.Print "Array("; i; ","; j; ") = "; vArray(i, j)
            Next
        Next

    '<EhFooter>
    Exit Sub
DumpArray2_Err:
    Debug.Print "GetSSLib.MDumper.DumpArray2:Error " & Err.Description
    
    '</EhFooter>
End Sub
Private Function ArrayBound(vArray As Variant, Optional GetUbound As Boolean = False) As Long
    '<EhHeader>
    On Error GoTo ArrayBound_Err
    '</EhHeader>
    If GetUbound Then
        ArrayBound = UBound(vArray)
    Else
        ArrayBound = LBound(vArray)
    End If
    '<EhFooter>
    Exit Function

ArrayBound_Err:
    If GetUbound Then
        ArrayBound = -1
    Else
        ArrayBound = -3
    End If
    Err.Clear
    
    '</EhFooter>
End Function


