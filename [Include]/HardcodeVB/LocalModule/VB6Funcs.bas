Attribute VB_Name = "MVB6Funcs"
Option Explicit

' New VB6 functions implemented for VB5

#If iVBVer <= 5 Then
Private Sub ErrRaise(e As Long)
    Dim sSource As String
    ' Raise standard Visual Basic error
    sSource = App.EXEName & ".VBError"
    Err.Raise e, sSource
End Sub

' A slow and inefficient version
Function InStrRev(sTarget As String, _
                          sFind As String, _
                          Optional ByVal iStart As Long = -1, _
                          Optional ByVal vcmCompare As VbCompareMethod) As Long
            
    ' Handle missing arguments
    If iStart = -1 Then iStart = Len(sTarget)
    
    ' Search backward
    Dim cFind As Long, i As Long, f As Long
    cFind = Len(sFind)
    For i = iStart - cFind + 1 To 1 Step -1
        If StrComp(Mid$(sTarget, i, cFind), sFind, vcmCompare) = 0 Then
            InStrRev = i
            Exit Function
        End If
    Next
End Function

' Filter for VB5
Function Filter(vInput As Variant, sMatch As String, _
                Optional fInclude As Boolean = True, _
                Optional Compare As VbCompareMethod = vbBinaryCompare _
                ) As Variant
    Dim asRet() As String, c As Long, i As Long, s As String
    On Error GoTo FilterResize
    For i = 0 To UBound(vInput)
        s = vInput(i)
        If InStr(1, s, sMatch, Compare) Then
            If fInclude Then
                asRet(c) = s
                c = c + 1
            End If
        Else
            If Not fInclude Then
                asRet(c) = s
                c = c + 1
            End If
        End If
    Next
    ReDim Preserve asRet(0 To c - 1)
    Filter = asRet
    Exit Function
    
FilterResize:
    Const cChunk As Long = 20
    If Err.Number = eeOutOfBounds Then
        ReDim Preserve asRet(0 To c + cChunk) As String
        Resume              ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Function

' Define the mighty Join function for VB5
Function Join(vaSourceArray As Variant, _
              Optional vsDelimiter As Variant) As String
    Dim i As Long
    If IsMissing(vsDelimiter) Then vsDelimiter = " "
    For i = LBound(vaSourceArray) To UBound(vaSourceArray) - 1
        Join = Join & vaSourceArray(i) & vsDelimiter
    Next
    Join = Join & vaSourceArray(i)
End Function

' VB5 users get a better Split than Split (except Compare parameter ignored)
Function Split(sExpression As String, _
               Optional Delimiters As Variant, _
               Optional Limit As Variant, _
               Optional Compare As VbCompareMethod) As Variant
    Dim sToken As String, avsRet() As Variant, c As Long
    Dim sDelimiters As String
    If IsMissing(Delimiters) Then Delimiters = sWhiteSpace
    sDelimiters = Delimiters
    If IsMissing(Limit) Then Limit = -1
    ' Ignore the user's pitiful request for case-sensitivity
    ' (actually left as an exercise for the reader)
    Compare = vbBinaryCompare
    ' Error trap to resize on overflow
    On Error GoTo SplitResize
    ' Break into tokens and put in an array
    sToken = GetToken(sExpression, sDelimiters)
    Do While sToken <> sEmpty
        If Limit <> -1 Then If c >= Limit Then Exit Do
        avsRet(c) = sToken
        c = c + 1
        sToken = GetToken(sEmpty, sDelimiters)
    Loop
    ' Size is an estimate, so resize to counted number of tokens
    If c Then ReDim Preserve avsRet(0 To c - 1) As Variant
    Split = avsRet
    Exit Function
    
SplitResize:
    ' Resize on overflow
    Const cChunk As Long = 20
    If Err.Number = eeOutOfBounds Then
        ReDim Preserve avsRet(0 To c + cChunk) As Variant
        Resume              ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Function

' Hey, kids! Write your very own VB6 functions for VB5! How about
' LoadPicture? Should be easy to fake (and do a better job) using
' the API LoadImage. See LoadAnyPicture for an example. Here are some
' others:
'
'   FormatCurrency (could be messy)
'   FormatDateTime (a challenge to make it international)
'   FormatNumber (you could do better than they did)
'   FormatPercent (similar to FormatNumber)
'   MonthName (easy at the first level, but try making it international)
'   Replace (interesting and useful project)
'   Round (fun challenge)
'   StrReverse (yes, but why?)
'   WeekdayName (same as MonthName)

#End If



