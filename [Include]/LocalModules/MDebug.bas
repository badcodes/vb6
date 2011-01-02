Attribute VB_Name = "MDebug"
Option Explicit
Public Sub ClearDebugLog()
On Error Resume Next
Kill App.Path & "\" & "DebugLog.txt"
End Sub
Public Sub DebugFPrint(ByRef sText As Variant, Optional ByRef filename As String = "", Optional ByRef CreateNew As Boolean = False)

    #If LDebug = 1 Then
        If filename = "" Then filename = "DebugLog.txt"
        Dim f As Integer
        f = FreeFile()
        If CreateNew = True Then
            Open App.Path & "\" & filename For Output As #f
        Else
            Open App.Path & "\" & filename For Append As #f
        End If
        Print #f, "[" & Date$ & "][" & Time$ & "]" & CStr(sText)
        Close #f
    #Else
        Debug.Print sText
    #End If

End Sub

Public Sub DebugVPrint(ByRef msg As Variant)

    #If LDebug = 1 Then
        MsgBox CStr(msg), vbInformation, "VDebugInfo"
    #Else
        Debug.Print CStr(msg)
    #End If

End Sub


Public Sub Debug_DumpArray(arr)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Debug.Print arr(i)
    Next
End Sub

