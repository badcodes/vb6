Attribute VB_Name = "TestCStringMap"
Option Explicit

Public Sub Test_CStringMap()
'    Dim a As New CStringMap
'    a.Map("A") = "AAAAAAAAAAAAAAAAAAAA"
'    a.Map("B") = "BBBBBBBBBBBBBBBBBBBB"
'
'    Load frmTask2
'    With frmTask2
'        .Init a
'        .SetMultiLine "A"
'        .SetFile "A"
'        .Show 1
'    End With
    
'
'
'    End With
'    frmMapEditor.Init a
'    frmMapEditor.SetMultiLine "A"
'    frmMapEditor.SetDirectory "A"
'    frmMapEditor.Show 1

'    Dim b As CStringLink
'    Set b = a.ToStringLink
'    Do Until b Is Nothing
'        Debug.Print b.Data;
'        If Not b.NextLink Is Nothing Then
'            Debug.Print "-->";
'        End If
'        Set b = b.NextLink
'    Loop
End Sub

Public Sub Test_KVEditor()
'    Dim a(0 To 3, 0 To 1) As String
'    a(0, 0) = "A"
'    a(0, 1) = "AAAAAAAAAAAAAAAAA"
'    a(1, 0) = "B"
'    a(1, 1) = "BBBBBBBBBBBBBBBBB"
'    a(2, 0) = "C"
'    a(2, 1) = "CCCCCCCCCCCCCCCCC"
'    a(3, 0) = "D"
'    a(3, 1) = "DDDDDDDDDDDDDDDDD"
'    'DumpArray a(), 2
'    Load frmKeyValueEditor
'    With frmKeyValueEditor
'        .Process a()
'        .SetKeyStyle "A", VCT_Combox
'        .SetKeyStyle "B", VCT_DIR
'        .SetKeyStyle "C", VCT_MultiLine
'        .SetKeyStyle "D", VCT_FILE
'        .Show 1
'    End With
    'frmKeyValueEditor.Process a()
    'DumpArray a, 2
    'DumpArray frmKeyValueEditor.Result, 2
End Sub

Public Sub TestTaskEditor()
    Load frmTask
    frmTask.Show 1
End Sub
