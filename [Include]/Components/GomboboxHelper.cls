VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GComboboxHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function Add_UniqueItem(ByRef cboBox As ComboBox, ByRef itemText As String, Optional ByVal cmpMethod As VbCompareMethod = vbBinaryCompare) As Boolean
        '<EhHeader>
        On Error GoTo Add_UniqueItem_Err
        '</EhHeader>

    Dim i As Long

100 Add_UniqueItem = False

102 If cboBox Is Nothing Then Exit Function

104 With cboBox

106     For i = 0 To .ListCount
108         If StrComp(.List(i), itemText, cmpMethod) = 0 Then Exit Function
        Next
        
110     .AddItem itemText

    End With

112 Add_UniqueItem = True

        '<EhFooter>
        Exit Function

Add_UniqueItem_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ssMDBQuery.MComboboxHelper.Add_UniqueItem " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

