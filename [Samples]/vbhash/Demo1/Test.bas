Attribute VB_Name = "Test"
Option Explicit

' error constants
Public Const errKeyNotFound As Long = &H1
Public Const errDuplicateKey As Long = &H2

Private AHash As CAHash
Private Collect As CCollect
     
Public Sub Raise(errno As Long, loc As String)
    Dim msg As String

    ' express error as text
    Select Case errno
    Case errKeyNotFound
        msg = "Key not found"
    Case errDuplicateKey
        msg = "Duplicate key"
    End Select

    Err.Raise errno + vbObjectError + &H200, App.EXEName & "." & loc, msg & "."
End Sub

Public Sub Test()
    Dim i As Long
    Dim rec() As Variant
    Dim key() As Variant
    Dim time(3) As Single
    Dim ln(3) As String
    Dim r As Variant
    Dim initialAlloc As Long
    Dim GrowthFactor As Single
    Dim StopWatch As CStopWatch

    initialAlloc = FDemo.ItemCount / 10
    If initialAlloc < 10 Then initialAlloc = 10
    GrowthFactor = 1.5
    Set StopWatch = New CStopWatch
    
    ReDim rec(1 To FDemo.ItemCount)
    ReDim key(1 To FDemo.ItemCount)

    ' create keys and records
    For i = 1 To FDemo.ItemCount
        key(i) = Rnd * FDemo.ItemCount
        rec(i) = Rnd * FDemo.ItemCount
    Next i

    Select Case FDemo.Method
    Case 0
        OHash.Init FDemo.Size
    Case 1
        Set AHash = New CAHash
        AHash.Init FDemo.Size, initialAlloc, GrowthFactor
    Case 2
        Set Collect = New CCollect
    End Select

    StopWatch.Reset
    For i = 1 To FDemo.ItemCount
        Select Case FDemo.Method
        Case 0
            OHash.Insert key(i), rec(i)
        Case 1
            AHash.Insert key(i), rec(i)
        Case 2
            Collect.Insert key(i), rec(i)
        End Select
    Next i
    time(0) = StopWatch.Elapsed / 1000!

    StopWatch.Reset
    For i = 1 To FDemo.ItemCount
        Select Case FDemo.Method
        Case 0
            r = OHash.Find(key(i))
        Case 1
            r = AHash.Find(key(i))
        Case 2
            r = Collect.Find(key(i))
        End Select
        If r <> rec(i) Then MsgBox "fail at " & i
    Next i
    time(1) = StopWatch.Elapsed / 1000!

    StopWatch.Reset
    For i = 1 To FDemo.ItemCount
        Select Case FDemo.Method
        Case 0
            OHash.Delete key(i)
        Case 1
            AHash.Delete key(i)
        Case 2
            Collect.Delete key(i)
        End Select
    Next i
    time(2) = StopWatch.Elapsed / 1000!
    
    ln(0) = vbCrLf & "Insert: " & vbTab & FormatNumber(time(0), 3) & " secs"
    ln(1) = vbCrLf & "Search: " & vbTab & FormatNumber(time(1), 3) & " secs"
    ln(2) = vbCrLf & "Delete: " & vbTab & FormatNumber(time(2), 3) & " secs"
    MsgBox "Execution Time" & ln(0) & ln(1) & ln(2)
    
    ' free memory
    Select Case FDemo.Method
    Case 0
        OHash.Term
    Case 1
        Set AHash = Nothing
    Case 2
        Set Collect = Nothing
    End Select

End Sub

