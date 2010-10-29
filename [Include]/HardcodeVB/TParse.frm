VERSION 5.00
Begin VB.Form FTParse 
   Caption         =   "Parse Tester"
   ClientHeight    =   5256
   ClientLeft      =   1092
   ClientTop       =   1500
   ClientWidth     =   6732
   Icon            =   "TParse.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5256
   ScaleWidth      =   6732
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox txtOut 
      Height          =   4776
      Left            =   1704
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   240
      Width           =   4800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   252
      TabIndex        =   1
      Top             =   1548
      Width           =   1215
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "&Parse"
      Height          =   495
      Left            =   252
      TabIndex        =   0
      Top             =   252
      Width           =   1215
   End
End
Attribute VB_Name = "FTParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    txtOut.Text = sEmpty
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub cmdParse_Click()

    Dim sCommand As String, sSeparator As String
    Dim ms As Currency, msEnd As Currency, s As String, sOut As String
    Dim sToken As String, i As Long, iMax As Long
    iMax = 500
    txtOut.Text = sEmpty
       
    sCommand = "This is a big, fat, bloody " & sCrLf
    sCommand = sCommand & sTab & "test of ""gross stuff""."
    sSeparator = " ,." & sTab & Chr(13) & Chr(10)
    
    sOut = sOut & "Input String: " & sCommand & sCrLf & sCrLf
    
    sToken = GetToken1(sCommand, sSeparator)
    Do While sToken <> sEmpty
        s = s & sToken & "#"
        sToken = GetToken1(sEmpty, sSeparator)
    Loop
    ProfileStart ms
    For i = 1 To iMax
        sToken = GetToken1(sCommand, sSeparator)
        Do While sToken <> sEmpty
            'BugMessage sToken
            sToken = GetToken1(sEmpty, sSeparator)
        Loop
    Next
    ProfileStop ms, msEnd
    BugMessage "GetToken1: " & msEnd
    sOut = sOut & "GetToken1: " & msEnd & sCrLf
    sOut = sOut & s & sCrLf
    s = sEmpty
    
    sToken = GetToken2(sCommand, sSeparator)
    Do While sToken <> sEmpty
        s = s & sToken & "#"
        sToken = GetToken2(sEmpty, sSeparator)
    Loop
    ProfileStart ms
    For i = 1 To iMax
        sToken = GetToken2(sCommand, sSeparator)
        Do While sToken <> sEmpty
            'BugMessage sToken
            sToken = GetToken2(sEmpty, sSeparator)
        Loop
    Next
    ProfileStop ms, msEnd
    BugMessage "GetToken2: " & msEnd
    sOut = sOut & "GetToken2: " & msEnd & sCrLf
    sOut = sOut & s & sCrLf
    s = sEmpty
   
    sToken = GetToken3(sCommand, sSeparator)
    Do While sToken <> sEmpty
        s = s & sToken & "#"
        sToken = GetToken3(sEmpty, sSeparator)
    Loop
    ProfileStart ms
    For i = 1 To iMax
        sToken = GetToken3(sCommand, sSeparator)
        Do While sToken <> sEmpty
            'BugMessage sToken
            sToken = GetToken3(sEmpty, sSeparator)
        Loop
    Next
    ProfileStop ms, msEnd
    BugMessage "GetToken3: " & msEnd
    sOut = sOut & "GetToken3: " & msEnd & sCrLf
    sOut = sOut & s & sCrLf
    s = sEmpty
    
    sToken = GetToken4(sCommand, sSeparator)
    Do While sToken <> sEmpty
        s = s & sToken & "#"
        sToken = GetToken4(sEmpty, sSeparator)
    Loop
    ProfileStart ms
    For i = 1 To iMax
        sToken = GetToken4(sCommand, sSeparator)
        Do While sToken <> sEmpty
            'BugMessage sToken
            sToken = GetToken4(sEmpty, sSeparator)
        Loop
    Next
    ProfileStop ms, msEnd
    BugMessage "GetToken4: " & msEnd
    sOut = sOut & "GetToken4: " & msEnd & sCrLf
    sOut = sOut & s & sCrLf
    s = sEmpty
        
    sToken = GetToken(sCommand, sSeparator)
    Do While sToken <> sEmpty
        s = s & sToken & "#"
        sToken = GetToken(sEmpty, sSeparator)
    Loop
    ProfileStart ms
    For i = 1 To iMax
        sToken = GetToken(sCommand, sSeparator)
        Do While sToken <> sEmpty
            'BugMessage sToken
            sToken = GetToken(sEmpty, sSeparator)
        Loop
    Next
    ProfileStop ms, msEnd
    BugMessage "GetToken: " & msEnd
    sOut = sOut & "GetToken: " & msEnd & sCrLf
    sOut = sOut & s & sCrLf
    s = sEmpty
    
    sToken = GetQToken(sCommand, sSeparator)
    Do While sToken <> sEmpty
        s = s & sToken & "#"
        sToken = GetQToken(sEmpty, sSeparator)
    Loop
    ProfileStart ms
    For i = 1 To iMax
        sToken = GetQToken(sCommand, sSeparator)
        Do While sToken <> sEmpty
            'BugMessage sToken
            sToken = GetQToken(sEmpty, sSeparator)
        Loop
    Next
    ProfileStop ms, msEnd
    BugMessage "GetQToken: " & msEnd
    sOut = sOut & "GetQToken: " & msEnd & sCrLf
    sOut = sOut & s & sCrLf
    s = sEmpty
    

    Dim av As Variant, iTok As Long, v As Variant
#If 0 Then
    av = Split(sCommand, sSeparator)
    For Each v In av
        s = s & v & "#"
    Next
#Else
    av = Split(sCommand, sSeparator)
    For iTok = 0 To UBound(av)
        s = s & av(iTok) & "#"
    Next
#End If
    ProfileStart ms
    av = Split(sCommand, sSeparator)
    For i = 1 To iMax
        For iTok = 0 To UBound(av)
            sToken = av(iTok)
        Next
    Next
    ProfileStop ms, msEnd
    BugMessage "Split: " & msEnd
    sOut = sOut & "Split: " & msEnd & sCrLf
    sOut = sOut & s & sCrLf
    s = sEmpty
    
    av = Splits(sCommand, sSeparator)
    For iTok = 0 To UBound(av)
        s = s & av(iTok) & "#"
    Next
    ProfileStart ms
    av = Splits(sCommand, sSeparator)
    For i = 1 To iMax
        For iTok = 0 To UBound(av)
            sToken = av(iTok)
        Next
    Next
    ProfileStop ms, msEnd
    BugMessage "Splits: " & msEnd
    sOut = sOut & "Splits: " & msEnd & sCrLf
    sOut = sOut & s & sCrLf
    s = sEmpty
    
    av = QSplits(sCommand, sSeparator)
    For iTok = 0 To UBound(av)
        s = s & av(iTok) & "#"
    Next
    ProfileStart ms
    av = QSplits(sCommand, sSeparator)
    For i = 1 To iMax
        For iTok = 0 To UBound(av)
            sToken = av(iTok)
        Next
    Next
    ProfileStop ms, msEnd
    BugMessage "QSplits: " & msEnd
    sOut = sOut & "QSplits: " & msEnd & sCrLf
    sOut = sOut & s & sCrLf
    s = sEmpty
    
    txtOut.Text = sOut

End Sub

