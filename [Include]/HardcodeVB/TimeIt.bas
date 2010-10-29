Attribute VB_Name = "MTimeIt"
Option Explicit

'$ Uses DEBUG.BAS UTILITY.BAS SORT.BAS

Private Declare Function GetVersionTmp Lib "kernel32" Alias "GetVersion" () As Long

#If fUseCpp Then
Private Declare Function LoWord5 Lib "vbutil32" Alias "LoWord" (ByVal dw As Long) As Integer
Private Declare Function HiWord5 Lib "vbutil32" Alias "HiWord" (ByVal dw As Long) As Integer
#End If

Private n As Long
Private iVar As Integer

Private Type TLoHiLong
    lo As Integer
    hi As Integer
End Type

Private Type TAllLong
    all As Long
End Type

Function LogicalAndVsNestedIf(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String, i As Integer, iIter As Long

    i = 21
    ProfileStart sec
    For iIter = 1 To cIter
        If i <= 20 And i >= 10 Then i = i + 1
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "If a And b Then: " & secOut & " sec" & sCrLf
    
    i = 21
    ProfileStart sec
    For iIter = 1 To cIter
        If i <= 20 Then If i >= 10 Then i = i + 1
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "If a Then If b Then: " & secOut & " sec" & sCrLf

    LogicalAndVsNestedIf = sMsg

End Function

Function ByValVsByRef(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String, n As Long
    Dim i As Integer, lng As Long, sng As Single, dbl As Double
    Dim v As Variant, s As String
    
    i = 5
    ProfileStart sec
    For n = 1 To cIter
        TestByValInt i
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Integer by value: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter
        TestByRefInt i
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Integer by reference: " & secOut & " sec" & sCrLf

    lng = 100000
    ProfileStart sec
    For n = 1 To cIter
        TestByValLong lng
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Long by value: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter
        TestByRefLong lng
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Long by reference: " & secOut & " sec" & sCrLf
    
    sng = 2.1
    ProfileStart sec
    For n = 1 To cIter
        TestByValSng sng
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Single by value: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter
        TestByRefSng sng
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Single by reference: " & secOut & " sec" & sCrLf

    dbl = 2.1
    ProfileStart sec
    For n = 1 To cIter
        TestByValDbl dbl
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Double by value: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter
        TestByRefDbl dbl
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Double by reference: " & secOut & " sec" & sCrLf

    v = CDbl(2.1)
    ProfileStart sec
    For n = 1 To cIter
        TestByValVar v
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Variant (Double) by value: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter
        TestByRefVar v
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Variant (Double) by reference: " & secOut & " sec" & sCrLf

    s = "Hardcore"
    ProfileStart sec
    For n = 1 To cIter
        TestByValStr s
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "String by value: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter
        TestByRefStr s
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "String by reference: " & secOut & " sec" & sCrLf

    ByValVsByRef = sMsg

End Function

Private Sub TestByValInt(ByVal iArg As Integer)
    Dim iVar As Integer
    iVar = iArg
End Sub

Private Sub TestByRefInt(ByRef iArg As Integer)
    Dim iVar As Integer
    iVar = iArg
End Sub

Private Sub TestByValLong(ByVal iArg As Long)
    Dim iVar As Long
    iVar = iArg
End Sub

Private Sub TestByRefLong(ByRef iArg As Long)
    Dim iVar As Long
    iVar = iArg
End Sub

Private Sub TestByValSng(ByVal rArg As Single)
    Dim rVar As Single
    rVar = rArg
End Sub

Private Sub TestByRefSng(ByRef rArg As Single)
    Dim rVar As Single
    rVar = rArg
End Sub

Private Sub TestByValDbl(ByVal rArg As Double)
    Dim rVar As Double
    rVar = rArg
End Sub

Private Sub TestByRefDbl(ByRef rArg As Double)
    Dim rVar As Double
    rVar = rArg
End Sub

Private Sub TestByValCy(ByVal cyArg As Currency)
    Dim cyVar As Currency
    cyVar = cyArg
End Sub

Private Sub TestByRefCy(ByRef cyArg As Currency)
    Dim cyVar As Currency
    cyVar = cyArg
End Sub

Private Sub TestByValVar(ByVal vArg As Variant)
    Dim vVar As Variant
    vVar = vArg
End Sub

Private Sub TestByRefVar(ByRef vArg As Variant)
    Dim vVar As Variant
    vVar = vArg
End Sub

Private Sub TestByValStr(ByVal sArg As String)
    Dim sVar As String
    sVar = sArg
End Sub

Private Sub TestByRefStr(ByRef sArg As String)
    Dim sVar As String
    sVar = sArg
End Sub

Function CompareTypeProcessing(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim i As Integer, v As Variant, l As Long
    Dim s As Single, d As Double, c As Currency
    Dim i2 As Integer, v2 As Variant, l2 As Long
    Dim s2 As Single, d2 As Double, c2 As Currency
    Dim ci As Integer, cv As Variant, cl As Long
    Dim cs As Single, cd As Double, cc As Currency
    ci = IIf(cIter < 32767&, cIter, 0)
    cv = cIter: cl = cIter: cs = cIter: cd = cIter: cc = cIter
    
    ProfileStart sec
    i = 1
    Do While i < ci
        i2 = 1
        Do While i2 < ci
            i2 = i2 + 1
        Loop
        i = i + 1
    Loop
    ProfileStop sec, secOut
    sMsg = "Integer: " & secOut & " sec" & sCrLf

    ProfileStart sec
    l = 1
    Do While l < cl
        l2 = 1
        Do While l2 < ci
            l2 = l2 + 1
        Loop
        l = l + 1
    Loop
    ProfileStop sec, secOut
    sMsg = sMsg & "Long: " & secOut & " sec" & sCrLf

    ProfileStart sec
    s = 1
    Do While s < cs
        s2 = 1
        Do While s2 < ci
            s2 = s2 + 1
        Loop
        s = s + 1
    Loop
    ProfileStop sec, secOut
    sMsg = sMsg & "Single: " & secOut & " sec" & sCrLf

    ProfileStart sec
    d = 1
    Do While d < cd
        d2 = 1
        Do While d2 < ci
            d2 = d2 + 1
        Loop
        d = d + 1
    Loop
    ProfileStop sec, secOut
    sMsg = sMsg & "Double: " & secOut & " sec" & sCrLf

    ProfileStart sec
    c = 1
    Do While c < cc
        c2 = 1
        Do While c2 < ci
            c2 = c2 + 1
        Loop
        c = c + 1
    Loop
    ProfileStop sec, secOut
    sMsg = sMsg & "Currency: " & secOut & " sec" & sCrLf

    ProfileStart sec
    v = 1
    Do While v < cv
        v2 = 1
        Do While v2 < ci
            v2 = v2 + 1
        Loop
        v = v + 1
    Loop
    ProfileStop sec, secOut
    sMsg = sMsg & "Variant: " & secOut & " sec" & sCrLf

    CompareTypeProcessing = sMsg

End Function

Function InlineVsFunction(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim i As Long, n As Long, d As Double

    i = 1
    ProfileStart sec
    
    For n = 1 To cIter
        i = i + 1
    Next

    ProfileStop sec, secOut
    sMsg = "Inline addition: " & secOut & " sec" & sCrLf

    i = 1
    ProfileStart sec
    For n = 1 To cIter
        i = AddEm(i, 5)
    Next

    ProfileStop sec, secOut
    sMsg = sMsg & "Function Addition: " & secOut & " sec" & sCrLf

    i = 1
    ProfileStart sec
    For n = 1 To cIter
        d = n ^ 5
    Next

    ProfileStop sec, secOut
    sMsg = sMsg & "Inline power: " & secOut & " sec" & sCrLf

    i = 1
    ProfileStart sec
    For n = 1 To cIter
        d = Power(n, 5)
    Next

    ProfileStop sec, secOut
    sMsg = sMsg & "Function power: " & secOut & " sec" & sCrLf

    InlineVsFunction = sMsg
                                    
End Function

Function FixedVsVariableString(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim c As Long, n As Long, s As String
    Dim sVariable As String
    Dim sFixed As String * 8
    sVariable = "Hardcore"
    sFixed = "Hardcore"

    ProfileStart sec
    For n = 1 To cIter
        sVariable = "Hardcore"
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign to variable-length string: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter
        sFixed = "Hardcore"
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign to fixed-length string: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For n = 1 To cIter
        s = Mid$(sVariable, 3, 2)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Pass variable-length string to Mid$: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter
        s = Mid$(sFixed, 3, 2)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Pass fixed-length string to Mid$: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter / 100
        s = s & sVariable
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Concatenate variable-length string: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For n = 1 To cIter / 100
        s = s & sFixed
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Concatenate fixed-length string: " & secOut & " sec" & sCrLf
    
    sVariable = "Time It"
    ProfileStart sec
    For n = 1 To cIter / 10
        c = FindWindow(sNullStr, sVariable)
    Next
    ProfileStop sec, secOut
    BugMessage "Window handle: " & Hex(c)
    sMsg = sMsg & "Pass variable-length string to API: " & secOut & " sec" & sCrLf

    Dim sFixed2 As String * 7
    sFixed = "Time It"
    ProfileStart sec
    For n = 1 To cIter / 10
        c = FindWindow(sNullStr, sFixed2)
    Next
    ProfileStop sec, secOut
    BugMessage "Window handle: " & Hex(c)
    sMsg = sMsg & "Pass fixed-length string to API: " & secOut & " sec" & sCrLf

    Dim sVariableBuf As String
    sVariableBuf = String(80, 0)
    ProfileStart sec
    For n = 1 To cIter / 10
        c = GetWindowText(FTimeIt.hWnd, sVariableBuf, 80)
    Next
    ProfileStop sec, secOut
    s = Left$(sVariableBuf, c)
    BugMessage "Window Text: " & s
    sMsg = sMsg & "Use variable-length string as API buffer: " & secOut & " sec" & sCrLf

    Dim sFixedBuf As String * 80
    ProfileStart sec
    For n = 1 To cIter / 10
        c = GetWindowText(FTimeIt.hWnd, sFixedBuf, 80)
    Next
    ProfileStop sec, secOut
    BugMessage "Window Text: " & s
    s = Left$(sFixedBuf, c)
    sMsg = sMsg & "Use fixed-length string as API buffer: " & secOut & " sec" & sCrLf

    FixedVsVariableString = sMsg
                                    
End Function

Private Function AddEm(ByVal i1 As Long, i2 As Long) As Long
    AddEm = i1 + i2
End Function

Private Function Power(ByVal i1 As Long, i2 As Long) As Double
    Power = i1 ^ i2
End Function

Function CompareLoWords(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim i As Long
    Dim f16Wrap As Integer, f16NoWrap As Integer
    Dim f32Wrap As Long, f32NoWrap As Long
    f32NoWrap = &H12345678
    f32Wrap = &HFEDCBA98
    BugMessage "Wrap: " & Hex$(f32Wrap) & "  NoWrap: " & Hex$(f32NoWrap)
    
    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = LoWord1(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "LoWord1 - AND positive: " & secOut & " sec" & sCrLf

#If 0 Then
    ' This causes overflow!
    ProfileStart sec
    For i = 1 To cIter
        f16Wrap = LoWord1(f32Wrap)
    Next
    ProfileStop sec, secOut
    sMsg = "LoWord1 - AND negative: " & secOut & " sec" & sCrLf
    BugMessage "LoWord1 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)
#End If
    sMsg = sMsg & "LoWord1 - AND negative: Overflow" & sCrLf
    BugMessage "LoWord1 Wrap: " & "FAIL" & "  NoWrap: " & Hex$(f16NoWrap)

    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = LoWord2(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "LoWord2 - AND positive after sign check: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        f16Wrap = LoWord2(f32Wrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "LoWord2 - OR negative after sign check: " & secOut & " sec" & sCrLf
    BugMessage "LoWord2 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)

    f16Wrap = LoWord3(f32Wrap)
    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = LoWord3(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "LoWord3 - Copy low word with LSet: " & secOut & " sec" & sCrLf
    BugMessage "LoWord3 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)

    f16Wrap = LoWord4(f32Wrap)
    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = LoWord4(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "LoWord4 - Copy low word with CopyMemory: " & secOut & " sec" & sCrLf
    BugMessage "LoWord4 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)
    
#If fUseCpp Then
    f16Wrap = LoWord5(f32Wrap)
    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = LoWord5(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "LoWord5 - AND low word in C++: " & secOut & " sec" & sCrLf
    BugMessage "LoWord5 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)
#End If

    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = HiWord1(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "HiWord1 - AND negative: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        f16Wrap = HiWord1(f32Wrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "HiWord1 - AND positive: " & secOut & " sec" & sCrLf
    BugMessage "HiWord1 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)

#If 0 Then
    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = HiWord2(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "HiWord2 - AND positive after sign check: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        f16Wrap = HiWord2(f32Wrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "HiWord2 - AND negative after sign check: " & secOut & " sec" & sCrLf
    BugMessage "HiWord2 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)
#End If

    f16Wrap = HiWord3(f32Wrap)
    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = HiWord3(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "HiWord3 - Copy high word with LSet: " & secOut & " sec" & sCrLf
    BugMessage "HiWord3 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)

    f16Wrap = HiWord4(f32Wrap)
    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = HiWord4(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "HiWord4 - Copy high word with CopyMemory: " & secOut & " sec" & sCrLf
    BugMessage "HiWord4 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)
    
#If fUseCpp Then
    f16Wrap = HiWord5(f32Wrap)
    ProfileStart sec
    For i = 1 To cIter
        f16NoWrap = HiWord5(f32NoWrap)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "HiWord5 - AND high word in C++: " & secOut & " sec" & sCrLf
    BugMessage "HiWord5 Wrap: " & Hex$(f16Wrap) & "  NoWrap: " & Hex$(f16NoWrap)
#End If

    CompareLoWords = sMsg

End Function

Function LoWord1(ByVal dw As Long) As Integer
    LoWord1 = dw And &HFFFF&
End Function

Function LoWord2(ByVal dw As Long) As Integer
    If dw And &H8000& Then
        LoWord2 = dw Or &HFFFF0000
    Else
        LoWord2 = dw And &HFFFF&
    End If
End Function

Function LoWord3(ByVal dw As Long) As Integer
    Dim lohi As TLoHiLong
    Dim all  As TAllLong
    all.all = dw
    LSet lohi = all
    LoWord3 = lohi.lo
End Function

Function LoWord4(ByVal dw As Long) As Integer
    Dim w As Integer
    CopyMemory w, dw, 2
    LoWord4 = w
    'CopyMemory LoWord4, dw, 2
End Function

Function HiWord1(ByVal dw As Long) As Integer
    HiWord1 = (dw And &HFFFF0000) \ 65536
End Function

Function HiWord2(ByVal dw As Long) As Integer
    HiWord2 = (dw And &HFFFF0000) \ 65536
End Function

Function HiWord3(ByVal dw As Long) As Integer
    Dim lohi As TLoHiLong
    Dim all  As TAllLong
    all.all = dw
    LSet lohi = all
    HiWord3 = lohi.hi
End Function

Function HiWord4(ByVal dw As Long) As Integer
    CopyMemory HiWord4, ByVal VarPtr(dw) + 2, 2
End Function

Function LShiftWordB(ByVal w As Integer, ByVal c As Integer) As Integer
    Dim dw As Long
    dw = w * (2 ^ c)
    If dw And &H8000& Then
        LShiftWordB = CInt(dw And &H7FFF&) Or &H8000
    Else
        LShiftWordB = dw And &HFFFF&
    End If
End Function

Function RShiftWordB(ByVal w As Integer, ByVal c As Integer) As Integer
    Dim dw As Long
    If c = 0 Then
        RShiftWordB = w
    Else
        dw = w And &HFFFF&
        dw = dw \ (2 ^ c)
        RShiftWordB = dw And &HFFFF&
    End If
End Function

Function IIfVsIfThen(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim i As Long, iRes As Integer
    Dim ix As Integer, iy As Integer
    ix = 40: iy = 50
    
    ProfileStart sec
    For i = 1 To cIter
        iRes = IIf(ix > iy, ix, iy)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "IIf: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        If ix > iy Then
            iRes = ix
        Else
            iRes = iy
        End If
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "If-Then-Else: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iRes = MyIIf(ix > iy, ix, iy)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "MyIIf (Variant): " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iRes = MyIIfInt(ix > iy, ix, iy)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "MyIIfInt: " & secOut & " sec" & sCrLf

    IIfVsIfThen = sMsg

End Function

Private Function MyIIf(vCondition As Variant, vTrue As Variant, vFalse As Variant) As Variant
    If vCondition Then
        MyIIf = vTrue
    Else
        MyIIf = vFalse
    End If
End Function

Private Function MyIIfInt(iCondition As Integer, iTrue As Integer, iFalse As Integer) As Integer
    If iCondition Then
        MyIIfInt = iTrue
    Else
        MyIIfInt = iFalse
    End If
End Function

Function DollarVsNone(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim i As Long, iPos As Integer, cInput As Integer
    Const sTest As String = "To VB or not to VB, that is the question..."
    cInput = Len(sTest)
    Dim sOut As String, vOut As Variant
    Dim sInput As String, vInput As Variant

    sInput = sTest
    ProfileStart sec
    For i = 1 To cIter
        For iPos = 1 To Len(sInput) - 1
            sOut = Mid$(sInput, iPos)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "s = Mid$(s, i): " & secOut & " sec" & sCrLf

    vInput = sTest
    ProfileStart sec
    For i = 1 To cIter
        For iPos = 1 To Len(sInput) - 1
            sOut = Mid$(vInput, iPos)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "s = Mid$(v, i): " & secOut & " sec" & sCrLf

    vInput = sTest
    ProfileStart sec
    For i = 1 To cIter
        For iPos = 1 To Len(sInput) - 1
            vOut = Mid$(vInput, iPos)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "v = Mid$(v, i): " & secOut & " sec" & sCrLf
    
    sInput = sTest
    ProfileStart sec
    For i = 1 To cIter
        For iPos = 1 To Len(sInput) - 1
            vOut = Mid$(sInput, iPos)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "v = Mid$(s, i): " & secOut & " sec" & sCrLf
    
    sInput = sTest
    ProfileStart sec
    For i = 1 To cIter
        For iPos = 1 To Len(sInput) - 1
            sOut = Mid(sInput, iPos)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "s = Mid(s, i): " & secOut & " sec" & sCrLf

    vInput = sTest
    ProfileStart sec
    For i = 1 To cIter
        For iPos = 1 To Len(sInput) - 1
            sOut = Mid(vInput, iPos)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "s = Mid(v, i): " & secOut & " sec" & sCrLf

    vInput = sTest
    ProfileStart sec
    For i = 1 To cIter
        For iPos = 1 To Len(sInput) - 1
            vOut = Mid(vInput, iPos)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "v = Mid(v, i): " & secOut & " sec" & sCrLf
    
    sInput = sTest
    ProfileStart sec
    For i = 1 To cIter
        For iPos = 1 To Len(sInput) - 1
            vOut = Mid(sInput, iPos)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "v = Mid(s, i): " & secOut & " sec" & sCrLf
    
    DollarVsNone = sMsg

End Function

Function EmptyVsQuotes(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim asTest(0 To 1) As String
    Dim f As Boolean, i As Long
    
    asTest(0) = "Test"
    asTest(1) = Empty
    
    ProfileStart sec
    For i = 1 To cIter
        f = (asTest(i Mod 2) = sEmpty)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "s = sEmpty (String constant): " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        f = (asTest(i Mod 2) = "")
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "s = """" (inline quotes): " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        f = (asTest(i Mod 2) = Empty)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "s = Empty (Variant constant): " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        f = (asTest(i Mod 2) = vbNullString)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "s = vbNullString (null pointer constant): " & secOut & " sec" & sCrLf

    EmptyVsQuotes = sMsg

End Function

Function WithWithout(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim i As Long, iTest As Integer
    Dim nc As CNull, rnc As CNull
    Set nc = New CNull
    
    ProfileStart sec
    For i = 1 To cIter
        nc.ProcProp = 5
        iTest = nc.ProcProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Qualified access, one read/write: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        With nc
            .ProcProp = 5
            iTest = .ProcProp
        End With
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "With access, one read/write: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        nc.ProcProp = 5
        iTest = nc.ProcProp
        nc.ProcProp = 6
        iTest = nc.ProcProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Qualified access, two read/write: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        With nc
            .ProcProp = 5
            iTest = .ProcProp
            .ProcProp = 6
            iTest = .ProcProp
        End With
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "With access, two read/write: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        nc.ProcProp = 5
        iTest = nc.ProcProp
        nc.ProcProp = 6
        iTest = nc.ProcProp
        nc.ProcProp = 7
        iTest = nc.ProcProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Qualified access, three read/write: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        With nc
            .ProcProp = 5
            iTest = .ProcProp
            .ProcProp = 6
            iTest = .ProcProp
            .ProcProp = 7
            iTest = .ProcProp
        End With
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "With access, three read/write: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        nc.ProcProp = 5
        iTest = nc.ProcProp
        nc.ProcProp = 6
        iTest = nc.ProcProp
        nc.ProcProp = 7
        iTest = nc.ProcProp
        nc.ProcProp = 8
        iTest = nc.ProcProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Qualified access, four read/write: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        With nc
            .ProcProp = 5
            iTest = .ProcProp
            .ProcProp = 6
            iTest = .ProcProp
            .ProcProp = 7
            iTest = .ProcProp
            .ProcProp = 8
            iTest = .ProcProp
        End With
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "With access, four read/write: " & secOut & " sec" & sCrLf
    
    WithWithout = sMsg

End Function

Function MethodVsProc(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim i As Long
    Dim iTest As Integer
    Dim nul As CNull, nulNew As New CNull, nulLate As Object
    
    Set nul = New CNull
    Set nulLate = New CNull
    
    ProfileStart sec
    For i = 1 To cIter
        iTest = nul.FuncMethod()
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call method function on object: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        iTest = nulNew.FuncMethod()
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call method function on New object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iTest = nulLate.FuncMethod()
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call method function on late-bound object: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        iTest = FuncProc()
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call private function: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        nul.SubMethod iTest
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Pass variable to method sub on object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        nulNew.SubMethod iTest
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Pass variable to method sub on New object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        nulLate.SubMethod iTest
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Pass variable to method sub on late-bound object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iTest = FuncProc()
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Pass variable to private sub: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        nul.ProcProp = 5
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign through property let on object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        nulNew.ProcProp = 5
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign through property let on New object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        nulLate.ProcProp = 5
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign through property let on late-bound object: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        PrivProp = 5
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign through private property let: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        iTest = nul.ProcProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign from property get on object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iTest = nulNew.ProcProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign from property get on New object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iTest = nulLate.ProcProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign from property get on late-bound object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iTest = PrivProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign from private property get: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        nul.PubProp = 5
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign to public property on object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        nulNew.PubProp = 5
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign to public property on New object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        nulLate.PubProp = 5
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign to public property on late-bound object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iVar = 5
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign to private variable: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        iTest = nul.PubProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign from public property on object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iTest = nulNew.PubProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign from public property on New object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iTest = nulLate.PubProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign from public property on late-bound object: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        iTest = iVar
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Assign from private variable: " & secOut & " sec" & sCrLf
    
    MethodVsProc = sMsg

End Function

Private Sub SubProc(i As Integer)
    i = 1
End Sub

Private Function FuncProc()
    FuncProc = 2
End Function

Private Property Let PrivProp(i As Integer)
    iVar = i
End Property

Private Property Get PrivProp() As Integer
    PrivProp = iVar
End Property

Function ForEachVsForI(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String, i As Long, nul As CNull
    Dim iTemp As Integer, s As String, v As Variant, ix As Integer
    Dim iMax As Long

    iMax = cIter
    cIter = 1
    
    Dim avNull() As Variant
    Dim aiNull() As Integer
    Dim asNull() As String
    Dim anulNull() As CNull
    ReDim avNull(1 To iMax) As Variant
    ReDim aiNull(1 To iMax) As Integer
    ReDim asNull(1 To iMax) As String
    ReDim anulNull(1 To iMax) As CNull
    Dim nNull As Collection
    Set nNull = New Collection
    ' You can register Scripting for VB5 and enable this
#If iVBVer > 5 Then
    Dim dicNull As Dictionary
    Set dicNull = New Dictionary
#End If
    Dim vecNull As CVector
    Set vecNull = New CVector
    Dim veciNull As CVectorInt
    Set veciNull = New CVectorInt
    Dim lstNull As New CList
    Set lstNull = New CList
    Dim walker As CListWalker
    Set walker = New CListWalker
    
    ' Create collection, arrays, vector, and list of iMax Integers
    For iTemp = 1 To iMax
        avNull(iTemp) = iTemp
        aiNull(iTemp) = iTemp
        nNull.Add iTemp
#If iVBVer > 5 Then
        dicNull.Add CStr(iTemp), iTemp
#End If
        vecNull(iTemp) = iTemp
        veciNull(iTemp) = iTemp
        lstNull.Add iTemp
    Next
    
    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To iMax
            ix = avNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on Variant Integer array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For Each v In avNull
            ix = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on Variant Integer array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To iMax
            ix = aiNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on Integer array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For Each v In aiNull
            ix = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on Integer array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To nNull.Count
            ix = nNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on Integer collection: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each v In nNull
            ix = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on Integer collection: " & secOut & " sec" & sCrLf

#If iVBVer > 5 Then
    ProfileStart sec
    For i = 1 To cIter
        For Each v In dicNull.Items
            ix = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on Integer dictionary: " & secOut & " sec" & sCrLf
#End If
    
    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To vecNull.Last
            ix = vecNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on Variant Integer vector: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each v In vecNull
            ix = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on Variant Integer vector: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To veciNull.Last
            ix = veciNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on Integer vector: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each v In veciNull
            ix = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on Integer vector: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To lstNull.Count
            ix = lstNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on Integer list: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each v In lstNull
            ix = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on Integer list: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        walker.Attach lstNull
        Do While walker.More
            ix = walker
        Loop
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Do While on Integer list: " & secOut & " sec" & sCrLf

#If 0 Then  ' Turned out to be uninteresting, but left for the curious
    ' Create a collection and an array of iMax strings
    Set nNull = Nothing
    Set nNull = New Collection
    vecNull.Last = vecNull.Chunk
    lstNull.Clear
    For iTemp = 1 To iMax
        avNull(iTemp) = "Item" & iTemp
        asNull(iTemp) = "Item" & iTemp
        nNull.Add "Item" & iTemp
        vecNull(i) = "Item" & iTemp
        lstNull.Add "Item" & iTemp
    Next
    
    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To iMax
            s = avNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on Variant String array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For Each v In avNull
            s = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on Variant String array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To iMax
            s = asNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on String array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For Each v In asNull
            s = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on String array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To nNull.Count
            s = nNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on String collection: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each v In nNull
            s = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on String collection: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To vecNull.Last
            s = vecNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on String vector: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each v In vecNull
            s = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on String vector: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To lstNull.Count
            s = lstNull(iTemp)
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on String list: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each v In lstNull
            s = v
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on String list: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        walker.Attach lstNull
        Do While walker.More
            s = walker
        Loop
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Do While on String list: " & secOut & " sec" & sCrLf
#End If

    ' Create a collection and an array of iMax objects
    Set nNull = Nothing
    Set nNull = New Collection
    vecNull.Last = vecNull.Chunk
    lstNull.Clear
    For iTemp = 1 To iMax
        Set nul = New CNull
        nul.PubProp = iTemp
        Set anulNull(iTemp) = nul
        Set avNull(iTemp) = nul
        Set anulNull(iTemp) = nul
        nNull.Add nul
        Set vecNull(iTemp) = nul
        lstNull.Add nul
    Next
    
    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To iMax
            ix = avNull(iTemp).PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on Variant object array: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each v In avNull
            ix = v.PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on Variant object array: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To iMax
            ix = anulNull(iTemp).PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on object array: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each v In anulNull
            ix = v.PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on object array: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To nNull.Count
            ix = nNull(iTemp).PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on object collection: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each nul In nNull
            ix = nul.PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on object collection: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To vecNull.Last
            ix = vecNull(iTemp).PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on object vector: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each nul In vecNull
            ix = nul.PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on object vector: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        For iTemp = 1 To lstNull.Count
            ix = lstNull(iTemp).PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For I on object list: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        For Each nul In lstNull
            ix = nul.PubProp
        Next
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "For Each on object list: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = 1 To cIter
        walker.Attach lstNull
        Do While walker.More
            ix = walker.Item.PubProp
        Loop
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Do While on object list: " & secOut & " sec" & sCrLf
    
    ForEachVsForI = sMsg

End Function

Function SortCollectVsArray(cIter As Long) As String
' Uncomment to step through and verify that everything works
'#Const fTestSorts = 1
#If fTestSorts Then
    cIter = 10
#End If
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim i As Long, iTemp As Long, c As Integer
    Dim av() As Variant
    Dim n As New Collection
    c = cIter
    ReDim av(1 To cIter) As Variant

    ProfileStart sec
    For i = 1 To c
        av(i) = i
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Fill array: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To c
        n.Add i
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Fill collection: " & secOut & " sec" & sCrLf
    ShowNA av(), n

    ProfileStart sec
    ShuffleArray av()
    ProfileStop sec, secOut
    sMsg = sMsg & "Shuffle array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    ShuffleCollection n
    ProfileStop sec, secOut
    sMsg = sMsg & "Shuffle collection: " & secOut & " sec" & sCrLf
    ShowNA av(), n

    ProfileStart sec
    SortArray av(), 1, c
    ProfileStop sec, secOut
    sMsg = sMsg & "Sort array: " & secOut & " sec" & sCrLf

    ProfileStart sec
    SortCollection n, 1, c
    ProfileStop sec, secOut
    sMsg = sMsg & "Sort collection: " & secOut & " sec" & sCrLf
    ShowNA av(), n

    Dim v As Variant, iPos As Long, f As Boolean
    v = Random(1, c)
    ProfileStart sec
    For i = 1 To 50
        f = BSearchArray(av(), v, iPos)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Search array 50 Times: " & secOut & " sec" & sCrLf
#If fTestSorts Then
    BugMessage "Array element " & v & IIf(f, sEmpty, "not ") & " found at " & iPos
#End If

    ProfileStart sec
    For i = 1 To 50
        f = BSearchCollection(n, v, iPos)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Search collection 50 Times: " & secOut & " sec" & sCrLf
#If fTestSorts Then
    BugMessage "Collection element " & v & IIf(f, sEmpty, "not ") & " found at " & iPos
#End If

    SortCollectVsArray = sMsg

End Function

#If fTestSorts Then
Sub ShowNA(av() As Variant, n As Collection)
    Dim s As String, i As Integer, v As Variant
    For i = LBound(av) To UBound(av)
        s = s & av(i) & " "
    Next
    BugMessage "Array: " & s & sCrLf
    s = sEmpty
    For Each v In n
        s = s & v & " "
    Next
    BugMessage "Collection: " & s & sCrLf
End Sub
#Else
Sub ShowNA(av() As Variant, n As Collection)
End Sub
#End If

Function AddCollect(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String
    Dim i As Long, iTemp As Long, cHalf As Long
    Dim n As Collection
    cHalf = cIter / 2
    
    Set n = New Collection
    ProfileStart sec
    For i = 1 To cHalf
        n.Add i
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Add first half to end of collection: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = cHalf + 1 To cIter
        n.Add i
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Add last half to end of collection: " & secOut & " sec" & sCrLf
    Set n = Nothing

    Set n = New Collection
    ProfileStart sec
    n.Add 1
    For i = 2 To cHalf
        n.Add i, , 1
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Add first half to start of collection: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = cHalf + 1 To cIter
        n.Add i, , 1
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Add last half to start of collection: " & secOut & " sec" & sCrLf
    Set n = Nothing

    Set n = New Collection
    ProfileStart sec
    n.Add 1
    For i = 2 To cHalf
        n.Add i, , i \ 2
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Add first half to middle of collection: " & secOut & " sec" & sCrLf

    ProfileStart sec
    For i = cHalf + 1 To cIter
        n.Add i, , i \ 2
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Add last half to middle of collection: " & secOut & " sec" & sCrLf
    Set n = Nothing

    AddCollect = sMsg

End Function

Function SortRecurseVsIterate(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String, i As Integer
    Dim aR() As Variant, ai() As Variant
    Dim helper As New CSortHelper
    
    ReDim aR(1 To cIter) As Variant
    ReDim ai(1 To cIter) As Variant
    esmMode = esmSortVal

    ' Fill all arrays
    For i = 1 To cIter
        aR(i) = i
        ai(i) = i
    Next
    
    ' Randomize with same random sequence for both
    Seed 33
    ShuffleArray aR(), helper
    Seed 33
    ShuffleArray ai(), helper
    
    ProfileStart sec
    SortArrayRec aR() ', helper, 1, CInt(c)
    ProfileStop sec, secOut
    sMsg = sMsg & "Sort recursively: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    SortArray ai() ', helper, 1, CInt(c)
    ProfileStop sec, secOut
    sMsg = sMsg & "Sort iteratively: " & secOut & " sec" & sCrLf
    
    SortRecurseVsIterate = sMsg

End Function

Function SortNameVsSortPoly(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim sMsg As String, v As Variant
    Dim i As Long, iTemp As Long, c As Integer
    Dim aN() As Variant, aP() As Variant
    Dim helper As New CSortHelper
    
    On Error Resume Next
    
    c = cIter
    ReDim aN(1 To cIter) As Variant
    ReDim aP(1 To cIter) As Variant
    esmMode = esmSortVal

    ' Fill all arrays
    For i = 1 To c
        aN(i) = i
        aP(i) = i
    Next
    
    ' Use same random sequence for both
    Rnd -1
    ProfileStart sec
    ShuffleArrayO aN
    ProfileStop sec, secOut
    sMsg = sMsg & "Shuffle with name-space hack: " & secOut & " sec" & sCrLf

    Rnd -1
    ProfileStart sec
    ShuffleArray aP(), helper
    ProfileStop sec, secOut
    sMsg = sMsg & "Shuffle with polymorphic hack: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    SortArrayO aN(), 1, c
    ProfileStop sec, secOut
    sMsg = sMsg & "Sort with name-space hack: " & secOut & " sec" & sCrLf

    ProfileStart sec
    SortArray aP(), 1, c, helper
    ProfileStop sec, secOut
    sMsg = sMsg & "Sort with polymorphic hack: " & secOut & " sec" & sCrLf
        
    Dim iPos As Long, f As Boolean
    v = Random(1, c)
    ProfileStart sec
    For i = 1 To 50
        f = BSearchArrayO(aN(), v, iPos)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Search 50 times with name-space hack: " & secOut & " sec" & sCrLf

    v = Random(1, c)
    ProfileStart sec
    For i = 1 To 50
        f = BSearchArray(aP(), v, iPos, helper)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Search 50 times with polymorphic hack: " & secOut & " sec" & sCrLf
    
    SortNameVsSortPoly = sMsg
    
End Function

Function CompareFindFiles(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim i As Long, sMsg As String, v As Variant
    Dim nFiles As Collection, vFile As Variant
    Dim sFind As String, sDir As String
    sFind = Environ$("COMSPEC")
    sDir = Left$(CurDir$, 3)
    ProfileStart sec
    For i = 1 To cIter
        Set nFiles = FindFilesDir(sFind, sDir)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Find files with Dir$: " & secOut & " sec" & sCrLf
    BugMessage "Files found by FindFilesDir: " & sCrLf
    For Each vFile In nFiles
        BugMessage vFile & sCrLf
    Next

    Set nFiles = Nothing
    ProfileStart sec
    For i = 1 To cIter
        Set nFiles = FindFiles(sFind, sDir)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & sCrLf & "Find files with FindFirstFile: " & secOut & " sec" & sCrLf
    BugMessage "Files found by FindFiles: " & sCrLf
    For Each vFile In nFiles
        BugMessage vFile & sCrLf
    Next
    CompareFindFiles = sMsg
    
End Function

Function DeclareVsTypeLib(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim i As Long, dw As Long
    Dim sMsg As String
    
    ProfileStart sec
    For i = 1 To cIter
        dw = GetVersionTmp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call Declare function: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        dw = GetVersion
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call type library function: " & secOut & " sec" & sCrLf
    DeclareVsTypeLib = sMsg
    
End Function

Function CompareExistFile(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim i As Long, f As Boolean, sMsg As String
    Dim sYes As String, sNo As String
    ' You're guaranteed to have this file
    sYes = Environ$("COMSPEC")
    ' You're guaranteed not to have this file
    sNo = GetTempFile
    On Error Resume Next
    Kill sNo
    On Error GoTo 0
    
    ProfileStart sec
    For i = 1 To cIter
        f = ExistFile(sYes)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "ExistFile (error trap) on file: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        f = ExistFile(sNo)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "ExistFile (error trap) on no file: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        f = ExistFileDir(sYes)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "ExistFileDir (API) on file: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        f = ExistFileDir(sNo)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "ExistFileDir (API) on no file: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        f = Exists(sYes)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Exists (Dir$) on file: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        f = Exists(sNo)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Exists (Dir$) on no file: " & secOut & " sec" & sCrLf
    
    CompareExistFile = sMsg
    
End Function

Function CompareFriendVsPublic(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim i As Long, dw As Long
    Dim sMsg As String
    Dim fvp As CNull
    
    Set fvp = New CNull
    
    ProfileStart sec
    For i = 1 To cIter
        dw = fvp.FriendProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call Friend property on class: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        dw = fvp.ProcProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call Public property on class: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        dw = FTimeIt.FriendProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call Friend property on form: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        dw = FTimeIt.ProcProp
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Call Public property on form: " & secOut & " sec" & sCrLf
    CompareFriendVsPublic = sMsg
    
End Function

' Efficient find files function
Function FindFiles(sTarget As String, _
                   Optional ByVal Start As String) As Collection

    ' Statics for less memory use in recursive procedure
    Static sName As String, sSpec As String, nFound As New Collection
    Static fd As WIN32_FIND_DATA, iLevel As Long
    Dim hFiles As Long, f As Boolean
    If Start = sEmpty Then Start = CurDir$
    ' Maintain level to ensure collection is cleared first time
    If iLevel = 0 Then
        Set nFound = Nothing
        Start = NormalizePath(Start)
    End If
    iLevel = iLevel + 1
    
    ' Find first file (get handle to find)
    hFiles = FindFirstFile(Start & "*.*", fd)
    f = (hFiles <> INVALID_HANDLE_VALUE)
    Do While f
        sName = ByteZToStr(fd.cFileName)
        ' Skip . and ..
        If Left$(sName, 1) <> "." Then
            sSpec = Start & sName
            If fd.dwFileAttributes And vbDirectory Then
                DoEvents
                ' Call recursively on each directory
                FindFiles sTarget, sSpec & "\"
            ElseIf StrComp(sName, sTarget, 1) = 0 Then ' Text comparison
                ' Store found files in collection
                nFound.Add sSpec
            End If
        End If
        ' Keep looping until no more files
        f = FindNextFile(hFiles, fd)
    Loop
    f = FindClose(hFiles)
    ' Return the matching files in collection
    Set FindFiles = nFound
    iLevel = iLevel - 1
End Function

' Inefficient find files function to show how bad Dir$ is
Function FindFilesDir(sTarget As String, _
                      Optional ByVal Start As String) As Collection

    ' Statics for less memory use in recursive procedure
    Static sName As String, sSpec As String, v As Variant
    Static nFound As New Collection, iLevel As Long
    Dim nDirNames As New Collection
    If Start = sEmpty Then Start = CurDir$
    If iLevel = 0 Then
        Set nFound = Nothing
        Start = NormalizePath(Start)
    End If
    iLevel = iLevel + 1
    
    ' Ignore errors so that VB invalid file name won't kill search
    ' (Basic fails on weird but legal Win32 names such as ??????)
    On Error Resume Next
    ' Get first file
    sName = Dir$(Start, vbDirectory)
    Do While sName <> sEmpty
        ' Skip . and ..
        If Left$(sName, 1) <> "." Then
            sSpec = Start & sName
            If GetAttr(sSpec) And vbDirectory Then
                ' Cache directory names in collection
                nDirNames.Add sName
            ElseIf StrComp(sName, sTarget, 1) = 0 Then ' Text comparison
                ' Store found files in collection
                nFound.Add sSpec
            End If
        End If
        ' Keep looping until no more files
        sName = Dir$()
    Loop

    ' Call recursively on each cached directory
    For Each v In nDirNames
        FindFilesDir sTarget, Start & v & "\"
    Next

    ' Return the count of matching files
    Set FindFilesDir = nFound
    iLevel = iLevel - 1

End Function

#If iVBVer > 5 Then
Function ArrayVsVariant(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim i As Long, dw As Long
    Dim asTmp() As String, avTmp() As Variant
    Dim sMsg As String
    
    ReDim asTmp(0 To 9) As String
    ReDim avTmp(0 To 9) As Variant
    For i = 0 To 9
        asTmp(i) = CStr(i)
        avTmp(i) = CStr(i)
    Next
    
    ProfileStart sec
    For i = 1 To cIter
        asTmp = AS2AS(asTmp)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Return array of strings from function: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        avTmp = AV2AV(avTmp)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Return array of variants from function: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        asTmp = VAS2VAS(asTmp)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Return variant with array of strings from function: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        avTmp = VAV2VAV(avTmp)
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Return variant with array of variants from function: " & secOut & " sec" & sCrLf
    
    ArrayVsVariant = sMsg
    
End Function

Function VAV2VAV(vavRetIn As Variant) As Variant
    Dim iFirst As Long, iLast As Long, i As Long, iRev As Long
    iFirst = LBound(vavRetIn)
    iLast = UBound(vavRetIn)
    ' Size return to the size of input array
    Dim avRet() As Variant
    ReDim avRet(iFirst To iLast) As Variant
    ' March from back to front of input, assigning front to back of output
    For i = iLast To iFirst Step -1
        avRet(iRev) = vavRetIn(i)
        iRev = iRev + 1
    Next
    VAV2VAV = avRet
End Function

Function VAS2VAS(vasIn As Variant) As Variant
    Dim iFirst As Long, iLast As Long, i As Long, iRev As Long
    iFirst = LBound(vasIn)
    iLast = UBound(vasIn)
    ' Size return to the size of input array
    Dim asRet() As String
    ReDim asRet(iFirst To iLast) As String
    ' March from back to front of input, assigning front to back of output
    For i = iLast To iFirst Step -1
        asRet(iRev) = vasIn(i)
        iRev = iRev + 1
    Next
    VAS2VAS = asRet
End Function

Function AS2AS(asIn() As String) As String()
    ' Any type of array allowed
    Dim iFirst As Long, iLast As Long, i As Long, iRev As Long
    iFirst = LBound(asIn)
    iLast = UBound(asIn)
    ' Size return to the size of input array
    Dim asRet() As String
    ReDim asRet(iFirst To iLast) As String
    ' March from back to front of input, assigning front to back of output
    For i = iLast To iFirst Step -1
        asRet(iRev) = asIn(i)
        iRev = iRev + 1
    Next
    AS2AS = asRet
End Function

Function AV2AV(avInt() As Variant) As Variant()
    ' Any type of array allowed
    Dim iFirst As Long, iLast As Long, i As Long, iRev As Long
    iFirst = LBound(avInt)
    iLast = UBound(avInt)
    ' Size return to the size of input array
    Dim avRet() As Variant
    ReDim avRet(iFirst To iLast) As Variant
    ' March from back to front of input, assigning front to back of output
    For i = iLast To iFirst Step -1
        avRet(iRev) = avInt(i)
        iRev = iRev + 1
    Next
    AV2AV = avRet
End Function
#End If

Function XorVsTmpSwap(cIter As Long) As String
    Dim sec As Currency, secOut As Currency
    Dim i1 As Integer, i2 As Integer
    Dim sMsg As String, i As Long
    i1 = 111: i2 = 222
    
    ProfileStart sec
    For i = 1 To cIter
        SwapIntegerTmp i1, i2
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Swap with temporary variable: " & secOut & " sec" & sCrLf
    
    ProfileStart sec
    For i = 1 To cIter
        SwapIntegerXor i1, i2
    Next
    ProfileStop sec, secOut
    sMsg = sMsg & "Swap with XOR operations: " & secOut & " sec" & sCrLf
    
    XorVsTmpSwap = sMsg
    
End Function

Sub SwapIntegerTmp(i1 As Integer, i2 As Integer)
    ' Obvious way uses temporary variable
    Dim iTmp As Integer
    i1 = iTmp
    i2 = i1
    i1 = iTmp
End Sub

Sub SwapIntegerXor(i1 As Integer, i2 As Integer)
    ' Tricky way uses Xor operator
    i1 = i1 Xor i2
    i2 = i1 Xor i2
    i1 = i2 Xor i1
End Sub



