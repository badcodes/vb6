VERSION 5.00
Begin VB.Form FStringTest 
   Caption         =   "Test String"
   ClientHeight    =   6540
   ClientLeft      =   1560
   ClientTop       =   3000
   ClientWidth     =   7524
   Icon            =   "TString.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   7524
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle 8"
      Height          =   372
      Index           =   7
      Left            =   240
      TabIndex        =   30
      Top             =   3480
      Width           =   972
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "String Swap"
      Height          =   372
      Left            =   6276
      TabIndex        =   29
      Top             =   4800
      Width           =   1056
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search4"
      Height          =   372
      Index           =   3
      Left            =   240
      TabIndex        =   28
      Top             =   5484
      Width           =   972
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Text"
      Height          =   372
      Left            =   6276
      TabIndex        =   27
      Top             =   5316
      Width           =   1056
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert3"
      Height          =   372
      Index           =   2
      Left            =   6276
      TabIndex        =   26
      Top             =   4164
      Width           =   1056
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert2"
      Height          =   372
      Index           =   1
      Left            =   6276
      TabIndex        =   25
      Top             =   3684
      Width           =   1056
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert1"
      Height          =   372
      Index           =   0
      Left            =   6276
      TabIndex        =   24
      Top             =   3204
      Width           =   1056
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search3"
      Height          =   372
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   5004
      Width           =   972
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search2"
      Height          =   372
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   4524
      Width           =   972
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Case Insensitive"
      Height          =   252
      Left            =   156
      TabIndex        =   21
      Top             =   6120
      Width           =   1512
   End
   Begin VB.CheckBox chkForward 
      Caption         =   "Forward"
      Height          =   252
      Left            =   156
      TabIndex        =   20
      Top             =   5892
      Value           =   1  'Checked
      Width           =   1116
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search1"
      Height          =   372
      Index           =   0
      Left            =   252
      TabIndex        =   19
      Top             =   4032
      Width           =   972
   End
   Begin VB.TextBox txtFind 
      Height          =   372
      Left            =   1464
      TabIndex        =   15
      Top             =   396
      Width           =   4632
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "Numeric"
      Height          =   372
      Left            =   6276
      TabIndex        =   14
      Top             =   1116
      Width           =   1056
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle 7"
      Height          =   372
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   3024
      Width           =   972
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle 6"
      Height          =   372
      Index           =   5
      Left            =   252
      TabIndex        =   12
      Top             =   2544
      Width           =   972
   End
   Begin VB.CommandButton cmdGerman 
      Caption         =   "German"
      Height          =   372
      Left            =   6276
      TabIndex        =   11
      Top             =   2556
      Width           =   1056
   End
   Begin VB.CommandButton cmdFrench 
      Caption         =   "French"
      Height          =   372
      Left            =   6276
      TabIndex        =   10
      Top             =   2076
      Width           =   1056
   End
   Begin VB.CommandButton cmdSpanish 
      Caption         =   "Spanish"
      Height          =   372
      Left            =   6276
      TabIndex        =   9
      Top             =   1596
      Width           =   1056
   End
   Begin VB.CommandButton cmdLong 
      Caption         =   "Long"
      Height          =   372
      Left            =   6276
      TabIndex        =   8
      Top             =   636
      Width           =   1056
   End
   Begin VB.CommandButton cmdShort 
      Caption         =   "Short"
      Height          =   372
      Left            =   6276
      TabIndex        =   7
      Top             =   156
      Width           =   1056
   End
   Begin VB.TextBox txtTime 
      Height          =   2532
      Left            =   1476
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3540
      Width           =   4644
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle 5"
      Height          =   372
      Index           =   4
      Left            =   252
      TabIndex        =   5
      Top             =   2064
      Width           =   972
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle 4"
      Height          =   372
      Index           =   3
      Left            =   252
      TabIndex        =   4
      Top             =   1584
      Width           =   972
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle 3"
      Height          =   372
      Index           =   2
      Left            =   252
      TabIndex        =   3
      Top             =   1104
      Width           =   972
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle 2"
      Height          =   372
      Index           =   1
      Left            =   252
      TabIndex        =   2
      Top             =   624
      Width           =   972
   End
   Begin VB.TextBox txtData 
      Height          =   2148
      Left            =   1488
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1116
      Width           =   4620
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle 1"
      Height          =   372
      Index           =   0
      Left            =   252
      TabIndex        =   0
      Top             =   144
      Width           =   972
   End
   Begin VB.Label lbl 
      Caption         =   "Output"
      Height          =   252
      Index           =   2
      Left            =   1476
      TabIndex        =   18
      Top             =   3300
      Width           =   852
   End
   Begin VB.Label lbl 
      Caption         =   "Target"
      Height          =   252
      Index           =   1
      Left            =   1464
      TabIndex        =   17
      Top             =   876
      Width           =   852
   End
   Begin VB.Label lbl 
      Caption         =   "Find"
      Height          =   252
      Index           =   0
      Left            =   1464
      TabIndex        =   16
      Top             =   156
      Width           =   852
   End
End
Attribute VB_Name = "FStringTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CharLowerBuffW Lib "USER32" ( _
    ByVal lpsz As Long, ByVal cchLength As Long) As Long
Private Declare Function CharUpperBuffW Lib "USER32" ( _
    ByVal lpsz As Long, ByVal cchLength As Long) As Long
Private Declare Function IsCharUpperW Lib "USER32" ( _
    ByVal cChar As Integer) As Long
Private Declare Function IsCharLowerW Lib "USER32" ( _
    ByVal cChar As Integer) As Long

Private cIter As Long, fHasShellWapi As Boolean
Private msStart As Long, msEnd As Long

Private Sub Form_Load()
    fHasShellWapi = HasShellWapi
    Show
    If Not fHasShellWapi Then
        cmdSearch(1).Enabled = False
        cmdSearch(2).Enabled = False
    End If
    
    ChDir App.Path
    ChDrive App.Path
    cmdShort_Click
    txtFind = "than"
    chkForward_Click
End Sub

Private Sub chkCase_Click()
    If fHasShellWapi Then
        cmdSearch(1).Enabled = Not ((chkForward.Value <> vbChecked) And _
                                    (chkCase.Value <> vbChecked))
    End If
End Sub

Private Sub chkForward_Click()
    cmdSearch(3).Enabled = (chkForward.Value <> vbChecked)
    If fHasShellWapi Then
        cmdSearch(1).Enabled = Not ((chkForward.Value <> vbChecked) And _
                                    (chkCase.Value <> vbChecked))
    End If
End Sub

Private Sub cmdClear_Click()
    txtTime = sEmpty
End Sub

Private Sub cmdSwap_Click()
    Dim s1 As String, s2 As String
    s1 = "Who's got the ball?"
    s2 = "What's the story?"
    txtData = "Before swap string 1: " & s1 & vbCrLf
    txtData = txtData & "Before swap string 2: " & s2 & vbCrLf
    SwapStrings s1, s2
    txtData = txtData & "After swap string 1: " & s1 & vbCrLf
    txtData = txtData & "After swap string 2: " & s2 & vbCrLf
End Sub

Private Sub cmdLong_Click()
    cIter = 50
    Dim ab() As Byte
    ab = LoadResData(101, "STRING")
    txtData = StrConv(ab, vbUnicode)
End Sub

Private Sub cmdShort_Click()
    cIter = 1000
    txtData = "Make things simpler than possible"
End Sub

Private Sub cmdNumber_Click()
    cIter = 1000
    txtData = "See 3,987,654,456,429 bugs"
End Sub

Private Sub cmdGerman_Click()
    cIter = 1000
    txtData = "Basic ist eine höhere Sprache"
End Sub

Private Sub cmdFrench_Click()
    cIter = 1000
    txtData = "Peut-être simplement pour embêter le monde"
End Sub

Private Sub cmdSpanish_Click()
    cIter = 1000
    txtData = "Sobre Visual Basic hay muchos libros, pero sólo conozco uno que pueda servirte para ir un paso más allá de lo aprendido aquí"
End Sub

Private Sub cmdToggle_Click(Index As Integer)
    Dim i As Long, s As String
    s = txtData
    MousePointer = vbHourglass
    Select Case Index + 1
    Case 1
        msStart = GetTickCount
        For i = 1 To cIter
            s = TCase1(s)
        Next
        msEnd = GetTickCount
        txtData = TCase1(s)
    Case 2
        msStart = GetTickCount
        For i = 1 To cIter
            s = TCase2(s)
        Next
        msEnd = GetTickCount
        txtData = TCase2(s)
    Case 3
        msStart = GetTickCount
        For i = 1 To cIter
            s = TCase3(s)
        Next
        msEnd = GetTickCount
        txtData = TCase3(s)
    Case 4
        msStart = GetTickCount
        For i = 1 To cIter
            s = TCase4(s)
        Next
        msEnd = GetTickCount
        txtData = TCase4(s)
    Case 5
        msStart = GetTickCount
        For i = 1 To cIter
            s = TCase5(s)
        Next
        msEnd = GetTickCount
        txtData = TCase5(s)
    Case 6
        Dim ab() As Byte
        ab = StrConv(s, vbFromUnicode)
        msStart = GetTickCount
        For i = 1 To cIter
            TCase6 ab
        Next
        msEnd = GetTickCount
        TCase6 ab
        txtData = StrConv(ab, vbUnicode)
    Case 7
        Dim abRes() As Byte
        ab = StrConv(s, vbFromUnicode)
        msStart = GetTickCount
        For i = 1 To cIter
            abRes = TCase7(ab)
        Next
        msEnd = GetTickCount
        abRes = TCase7(ab)
        txtData = StrConv(abRes, vbUnicode)
    Case 8
        ab = StrConv(s, vbFromUnicode)
        msStart = GetTickCount
        For i = 1 To cIter
            TCase8 ab
        Next
        msEnd = GetTickCount
        TCase8 ab
        txtData = StrConv(ab, vbUnicode)
    End Select
    MousePointer = vbDefault
    txtTime = txtTime & "TCase" & Index + 1 & " time: " & msEnd - msStart & vbCrLf
    txtTime.SelStart = Len(txtTime)
End Sub

Private Sub cmdInsert_Click(Index As Integer)
    Dim i As Long, s As String, sNew As String
    s = "What's the deal"
    sNew = "big "
    cIter = 1500
    MousePointer = vbHourglass
    Select Case Index + 1
    Case 1
        msStart = GetTickCount
        For i = 1 To cIter
            InsertString1 sNew, s, 11
        Next
        msEnd = GetTickCount
        txtData = s
    Case 2
        msStart = GetTickCount
        For i = 1 To cIter
            InsertString2 sNew, s, 11
        Next
        msEnd = GetTickCount
        txtData = s
    Case 3
        msStart = GetTickCount
        For i = 1 To cIter
            InsertString3 sNew, s, 11
        Next
        msEnd = GetTickCount
        txtData = s
    End Select
    MousePointer = vbDefault
    txtTime = txtTime & "Insert" & Index + 1 & " time: " & msEnd - msStart & vbCrLf
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    Dim i As Long, iPos As Long
    Dim s As String, sFind As String, sFunc As String
    cIter = 5000
    s = txtData
    sFind = txtFind
    MousePointer = vbHourglass
    Select Case Index + 1
    Case 1
        Dim ordFind As Integer
        If chkCase = vbChecked Then
            ordFind = vbTextCompare
            sFunc = "Case sensitive "
        Else
            ordFind = vbBinaryCompare
            sFunc = "Case insensitive "
        End If
        If chkForward <> vbChecked Then
            iPos = -1
            sFunc = sFunc & "InStrRev"
            msStart = GetTickCount
            For i = 1 To cIter
                iPos = InStrRev(s, sFind, , ordFind)
            Next
            msEnd = GetTickCount
        Else
            sFunc = sFunc & "InStr"
            msStart = GetTickCount
            For i = 1 To cIter
                iPos = InStr(1, s, sFind, ordFind)
            Next
            msEnd = GetTickCount
        End If
        
    Case 2
        If chkForward = vbChecked Then
            If chkCase = vbChecked Then
                sFunc = "FindStrI"
                msStart = GetTickCount
                For i = 1 To cIter
                    iPos = FindStrI(s, sFind)
                Next
                msEnd = GetTickCount
            Else
                sFunc = "FindStr"
                msStart = GetTickCount
                For i = 1 To cIter
                    iPos = FindStr(s, sFind)
                Next
                msEnd = GetTickCount
            End If
        Else
            If chkCase = vbChecked Then
                sFunc = "FindStrRI"
                msStart = GetTickCount
                For i = 1 To cIter
                    iPos = FindStrRI(s, sFind)
                Next
                msEnd = GetTickCount
            End If
        End If
    Case 3
        If chkForward = vbChecked Then
            If chkCase = vbChecked Then
                sFunc = "FindStringI"
                msStart = GetTickCount
                For i = 1 To cIter
                    iPos = FindStringI(s, sFind)
                Next
                msEnd = GetTickCount
            Else
                sFunc = "FindString"
                msStart = GetTickCount
                For i = 1 To cIter
                    iPos = FindString(s, sFind)
                Next
                msEnd = GetTickCount
            End If
        Else
            If chkCase = vbChecked Then
                sFunc = "FindStringRI"
                msStart = GetTickCount
                For i = 1 To cIter
                    iPos = FindStringRI(s, sFind)
                Next
                msEnd = GetTickCount
            Else
                sFunc = "FindStringR"
                msStart = GetTickCount
                For i = 1 To cIter
                    iPos = FindStringR(s, sFind)
                Next
                msEnd = GetTickCount
            End If
        End If
    Case 4
        If chkCase = vbChecked Then
            ordFind = vbTextCompare
            sFunc = "Case sensitive "
        Else
            ordFind = vbBinaryCompare
            sFunc = "Case insensitive "
        End If
        If chkForward <> vbChecked Then
            iPos = -1
            sFunc = sFunc & "InStrR"
            msStart = GetTickCount
            For i = 1 To cIter
                iPos = InStrR(s, sFind, ordFind)
            Next
            msEnd = GetTickCount
        End If
    End Select
    MousePointer = vbDefault
    Select Case iPos
    Case -1
        ' String already assigned
    Case 0
        txtTime = txtTime & sFunc & " time: " & msEnd - msStart & vbCrLf
        txtTime = txtTime & "String '" & sFind & "' not found" & vbCrLf
    Case Else
        txtTime = txtTime & sFunc & " time: " & msEnd - msStart & vbCrLf
        txtTime = txtTime & "String '" & sFind & "' found at position " & iPos & vbCrLf
    End Select

End Sub

Function TCase1(ByVal s As String) As String
    Dim i As Long, ch As Integer
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        Select Case ch
        Case 65 To 90       ' A to Z
            ch = ch + 32
        Case 97 To 122      ' a to z
            ch = ch - 32
        ' Ignore non-characters
        End Select
        Mid$(s, i, 1) = Chr$(ch)
    Next
    TCase1 = s
End Function

Function TCase2(ByVal s As String) As String
    Dim i As Long, sCh As String
    For i = 1 To Len(s)
        sCh = Mid$(s, i, 1)
        If IsCharUpper(Asc(sCh)) Then
            sCh = LCase$(sCh)
        Else
            sCh = UCase$(sCh)
        End If
        Mid$(s, i, 1) = sCh
    Next
    TCase2 = s
End Function

Function TCase3(ByVal s As String) As String
    Dim i As Long, ch As Integer
    For i = 0 To Len(s) - 1
        CopyMemory ch, ByVal StrPtr(s) + (i * 2), 2
        If IsCharUpperW(ch) Then
            CharLowerBuffW StrPtr(s) + (i * 2), 1
        Else
            CharUpperBuffW StrPtr(s) + (i * 2), 1
        End If
    Next
    TCase3 = s
End Function

Function TCase4(ByVal s As String) As String
    Dim i As Long, ch As Byte, c As Long
    For i = 0 To Len(s) - 1
        CopyMemory ch, ByVal StrPtr(s) + (i * 2), 1
        If IsCharUpper(ch) Then
            CharLowerBuffPtr StrPtr(s) + (i * 2), 1
        Else
            CharUpperBuffPtr StrPtr(s) + (i * 2), 1
        End If
    Next
    TCase4 = s
End Function

Function TCase5(s As String) As String
    Dim i As Long, ab() As Byte
    ab = StrConv(s, vbFromUnicode)
    For i = 0 To Len(s) - 1
        If IsCharUpper(ab(i)) Then
            CharLowerBuffB ab(i), 1
        Else
            CharUpperBuffB ab(i), 1
        End If
    Next
    TCase5 = StrConv(ab, vbUnicode)
End Function

Sub TCase6(ab() As Byte)
    Dim i As Long
    For i = LBound(ab) To UBound(ab)
        If IsCharUpper(ab(i)) Then
            CharLowerBuffB ab(i), 1
        Else
            CharUpperBuffB ab(i), 1
        End If
    Next
End Sub

#If iVBVer > 5 Then
Function TCase7(ab() As Byte) As Byte()
    Dim i As Long, abRet() As Byte
    ' Copy the input array to a temporary value and process
    abRet = ab
    For i = LBound(abRet) To UBound(abRet)
        If IsCharUpper(abRet(i)) Then
            CharLowerBuffB abRet(i), 1
        Else
            CharUpperBuffB abRet(i), 1
        End If
    Next
    ' Return temporary
    TCase7 = abRet
End Function
#Else
' VB5 version must fake it with Variant return
Function TCase7(ab() As Byte) As Variant
    Dim i As Long, abRet() As Byte
    ' Copy the input array to a temporary value and process
    abRet = ab
    For i = LBound(abRet) To UBound(abRet)
        If IsCharUpper(abRet(i)) Then
            CharLowerBuffB abRet(i), 1
        Else
            CharUpperBuffB abRet(i), 1
        End If
    Next
    ' Return temporary
    TCase7 = abRet
End Function
#End If

Sub TCase8(ab() As Byte)
    Dim i As Long
    For i = LBound(ab) To UBound(ab)
        If IsCharUpper(ab(i)) Then
            CharLowerBuffB ab(i), 1
        ElseIf IsCharLower(ab(i)) Then
            CharUpperBuffB ab(i), 1
        End If
    Next
End Sub

Function FindStr(sTarget As String, sFind As String, _
                 Optional ByVal iPos As Long = 0) As Long
    ' FindStr = 0
    Dim pTarget As Long, pRet As Long
    ' Default position is start of string
    If iPos = 0 Then iPos = 1
    ' Turn string into pointer
    pTarget = StrPtr(sTarget)
    ' Get a pointer to the first match (if any)
    pRet = StrStr(pTarget + ((iPos - 1) * 2), StrPtr(sFind))
    ' If match, convert pointer to a string index
    If pRet Then FindStr = ((pRet - pTarget) / 2) + 1
End Function

Function FindStrI(sTarget As String, sFind As String, _
                  Optional ByVal iPos As Long = 0) As Long
    ' FindStrI = 0
    Dim pTarget As Long, pRet As Long
    ' Default position is start of string
    If iPos = 0 Then iPos = 1
    ' Turn string into pointer
    pTarget = StrPtr(sTarget)
    pRet = StrStrI(pTarget + ((iPos - 1) * 2), StrPtr(sFind))
    If pRet Then FindStrI = ((pRet - pTarget) / 2) + 1
End Function

' No FindStrR because no StrRStr

Function FindStrRI(sTarget As String, sFind As String, _
                   Optional ByVal iPos As Long = 0) As Long
    ' FindStrRI = 0
    Dim pTarget As Long, pRet As Long
    ' Default position is end of string
    If iPos = 0 Then iPos = Len(sTarget)
    ' Turn string into pointer
    pTarget = StrPtr(sTarget)
    ' Get a pointer to the last match (if any)
    pRet = StrRStrI(pTarget, (pTarget + ((iPos - 1) * 2)), _
                    StrPtr(sFind))
    ' If match, convert pointer to a string index
    If pRet Then FindStrRI = ((pRet - pTarget) / 2) + 1
End Function

Function FindString(sTarget As String, sFind As String, _
                    Optional ByVal iPos As Long = 0) As Long
    Dim pTarget As Long, pLast As Long, pFind As Long, pCur As Long
    Dim ch As Integer, cFind As Long
    ' FindString = 0
    If iPos = 0 Then iPos = 1
    ' Get start of string for position calculation
    pTarget = StrPtr(sTarget)
    ' Calculate pointer for start of search
    pCur = pTarget + ((iPos - 1) * 2)
    ' Get pointer to second character of find string
    pFind = StrPtr(sFind) + 2
    ' Calculate some helper values
    cFind = Len(sFind)
    pLast = pCur + ((Len(sTarget) - cFind) * 2)
    ' Get first character
    CopyMemory ch, ByVal pFind - 2, 2
    ' Special case for finding single character
    If cFind = 1 Then
        ' Find the first character of find string
        pCur = StrChr(pCur, ch)
        If pCur <> 0 Then
            ' Calculate position of found string
            FindString = ((pCur - pTarget) / 2) + 1
        End If
        Exit Function
    End If
    ' Normal case of multi-character search
    Do
        ' Find the first character of find string
        pCur = StrChr(pCur, ch)
        ' If first character not found, string can't be found
        If pCur = 0 Then Exit Function
        ' If not enough characters left for a match, don't check
        If pCur > pLast Then Exit Function
        ' Skip to second character and compare rest of string
        If StrCmpN(pCur + 2, pFind, cFind - 1) = 0 Then
            ' Calculate position of found string
            FindString = ((pCur - pTarget) / 2) + 1
            Exit Function
        End If
        pCur = pCur + 2
    Loop
    
End Function

Function FindStringI(sTarget As String, sFind As String, _
                     Optional ByVal iPos As Long = 0) As Long
    Dim pTarget As Long, pLast As Long, pFind As Long, pCur As Long
    Dim ch As Integer, cFind As Long
    ' FindStringI = 0
    If iPos = 0 Then iPos = 1
    ' Get start of string for position calculation
    pTarget = StrPtr(sTarget)
    ' Calculate pointer for start of search
    pCur = pTarget + ((iPos - 1) * 2)
    ' Get pointer to second character of find string
    pFind = StrPtr(sFind) + 2
    ' Calculate some helper values
    cFind = Len(sFind)
    pLast = pCur + ((Len(sTarget) - cFind) * 2)
    ' Get first character
    CopyMemory ch, ByVal pFind - 2, 2
    ' Special case for finding single character
    If cFind = 1 Then
        ' Find the first character of find string
        pCur = StrChrI(pCur, ch)
        If pCur <> 0 Then
            ' Calculate position of found string
            FindStringI = ((pCur - pTarget) / 2) + 1
        End If
        Exit Function
    End If
    ' Normal case of multi-character search
    Do
        ' Find the first character of find string
        pCur = StrChrI(pCur, ch)
        ' If first character not found, string can't be found
        If pCur = 0 Then Exit Function
        ' If not enough characters left for a match, don't check
        If pCur > pLast Then Exit Function
        ' Skip to second character and compare rest of string
        If StrCmpNI(pCur + 2, pFind, cFind - 1) = 0 Then
            ' Calculate position of found string
            FindStringI = ((pCur - pTarget) / 2) + 1
            Exit Function
        End If
        pCur = pCur + 2
    Loop
End Function

Function FindStringRI(sTarget As String, sFind As String, _
                      Optional ByVal iPos As Long = 0) As Long
    Dim pTarget As Long, pCur As Long, pLast As Long, pFind As Long
    Dim ch As Integer, cFind As Long
    ' FindStringR = 0
    ' Calculate some helper values
    cFind = Len(sFind)
    pTarget = StrPtr(sTarget)
    pLast = pTarget + ((cFind - 1) * 2)
    If iPos = 0 Then iPos = Len(sTarget)
    ' Get last possible match for position calculation
    pCur = pTarget + ((iPos - 1) * 2) - ((cFind - 1) * 2)
    ' Get pointer to second character of find string
    pFind = StrPtr(sFind) + 2
    ' Get first character
    CopyMemory ch, ByVal pFind - 2, 2
    ' Special case for finding single character
    If cFind = 1 Then
        ' Find the first character of find string
        pCur = StrRChrI(pTarget, pCur, ch)
        If pCur <> 0 Then
            ' Calculate position of found string
            FindStringRI = ((pCur - pTarget) / 2) + 1
        End If
        Exit Function
    End If
    ' Normal case of multi-character search
    Do
        ' Find the first character of find string
        pCur = StrRChrI(pTarget, pCur, ch)
        ' If first character not found, string can't be found
        If pCur = 0 Then Exit Function
        ' If not enough characters left for a match, don't check
        If pCur < pLast Then Exit Function
        ' Skip to second character and compare rest of string
        If StrCmpNI(pCur + 2, pFind, cFind - 1) = 0 Then
            ' Calculate position of found string
            FindStringRI = ((pCur - pTarget) / 2) + 1
            Exit Function
        End If
        pCur = pCur - 2
    Loop
End Function

Function FindStringR(sTarget As String, sFind As String, _
                     Optional ByVal iPos As Long = 0) As Long
    Dim pTarget As Long, pCur As Long, pLast As Long, pFind As Long
    Dim ch As Integer, cFind As Long
    ' FindStringR = 0
    ' Calculate some helper values
    cFind = Len(sFind)
    pTarget = StrPtr(sTarget)
    pLast = pTarget + ((cFind - 1) * 2)
    If iPos = 0 Then iPos = Len(sTarget)
    ' Get last possible match for position calculation
    pCur = pTarget + ((iPos - 1) * 2) - ((cFind - 1) * 2)
    ' Get pointer to second character of find string
    pFind = StrPtr(sFind) + 2
    ' Get first character
    CopyMemory ch, ByVal pFind - 2, 2
    ' Special case for finding single character
    If cFind = 1 Then
        ' Find the first character of find string
        pCur = StrRChr(pTarget, pCur, ch)
        If pCur <> 0 Then
            ' Calculate position of found string
            FindStringR = ((pCur - pTarget) / 2) + 1
        End If
        Exit Function
    End If
    ' Normal case of multi-character search
    Do
        ' Find the first character of find string
        pCur = StrRChr(pTarget, pCur, ch)
        ' If first character not found, string can't be found
        If pCur = 0 Then Exit Function
        ' If not enough characters left for a match, don't check
        If pCur < pLast Then Exit Function
        ' Skip to second character and compare rest of string
        If StrCmpN(pCur + 2, pFind, cFind - 1) = 0 Then
            ' Calculate position of found string
            FindStringR = ((pCur - pTarget) / 2) + 1
            Exit Function
        End If
        pCur = pCur - 2
    Loop
End Function

Function InStrR(Optional vStart As Variant, _
                Optional vTarget As Variant, _
                Optional vFind As Variant, _
                Optional vCompare As Variant) As Long
    If IsMissing(vStart) Then Err.Raise 5000, "MStrTool", "Missing parameter"
    
    ' Handle missing arguments
    Dim iStart As Long, sTarget As String
    Dim sFind As String, ordCompare As Long
    If VarType(vStart) = vbString Then
        If IsMissing(vTarget) Then Err.Raise 5000, "MStrTool", "Missing parameter"
        sTarget = vStart
        sFind = vTarget
        iStart = Len(sTarget)
        If IsMissing(vFind) Then
            ordCompare = vbBinaryCompare
        Else
            ordCompare = vFind
        End If
    Else
        If IsMissing(vTarget) Or IsMissing(vFind) Then
            Err.Raise 5000, "MStrTool", "Missing parameter"
        End If
        sTarget = vTarget
        sFind = vFind
        iStart = vStart
        If IsMissing(vCompare) Then
            ordCompare = vbBinaryCompare
        Else
            ordCompare = vCompare
        End If
    End If
    
    ' Search backward
    Dim cFind As Long, i As Long, f As Long
    cFind = Len(sFind)
    For i = iStart - cFind + 1 To 1 Step -1
        If StrComp(Mid$(sTarget, i, cFind), sFind, ordCompare) = 0 Then
            InStrR = i
            Exit Function
        End If
    Next
End Function

Sub InsertString1(sNew As String, sTarget As String, i As Long)
    sTarget = Left$(sTarget, i) & sNew & Mid$(sTarget, i + 1)
End Sub

Sub InsertString2(sNew As String, sTarget As String, i As Long)
    Dim s As String, cNew As Long
    cNew = Len(sNew)
    s = String$(cNew + Len(sTarget), 0)
    Mid$(s, 1, i) = sTarget
    Mid$(s, i + 1) = sNew
    Mid$(s, i + 1 + cNew) = Mid$(sTarget, i + 1)
    sTarget = s
End Sub

Sub InsertString3(sNew As String, sTarget As String, i As Long)
    Dim s As String, cNew As Long, cTarget As Long
    Dim p As Long, pTarget As Long
    cNew = Len(sNew)
    cTarget = Len(sTarget)
    s = String$(cNew + cTarget, 0)
    p = StrPtr(s)
    pTarget = StrPtr(sTarget)
    CopyMemory ByVal p, ByVal pTarget, i * 2
    CopyMemory ByVal p + (i * 2), ByVal StrPtr(sNew), cNew * 2
    CopyMemory ByVal p + ((i + cNew) * 2), _
               ByVal pTarget + (i * 2), (cTarget - i) * 2
    sTarget = s
End Sub


