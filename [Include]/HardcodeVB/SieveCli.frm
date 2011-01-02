VERSION 5.00
Object = "{04C14C62-BFAC-11D0-B253-00AA005754FD}#1.0#0"; "SieveBasCtlN.ocx"
Object = "{B4A64CE4-D292-11D0-B253-00AA005754FD}#1.0#0"; "SieveBasCtlP.ocx"
Begin VB.Form FSieveClient 
   Caption         =   "Sieve of Eratosthenes Client"
   ClientHeight    =   2712
   ClientLeft      =   3000
   ClientTop       =   1728
   ClientWidth     =   5616
   Icon            =   "SieveCli.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2712
   ScaleWidth      =   5616
   Begin SieveBasCtlN.XSieveN sieveCtlN 
      Height          =   540
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin SieveBasCtlP.XSieveP sieveCtlP 
      Height          =   540
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.ComboBox cboServer 
      Height          =   288
      ItemData        =   "SieveCli.frx":0CFA
      Left            =   1596
      List            =   "SieveCli.frx":0D1F
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2040
      Width           =   2760
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Get All"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   2388
      Width           =   828
   End
   Begin VB.CheckBox chkLate 
      Caption         =   "Late Bind"
      Height          =   255
      Left            =   2556
      TabIndex        =   12
      Top             =   2388
      Width           =   975
   End
   Begin VB.ListBox lstOutput 
      Height          =   1968
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   972
   End
   Begin VB.CheckBox chkDisplay 
      Caption         =   "Display"
      Height          =   255
      Left            =   1596
      TabIndex        =   10
      Top             =   2388
      Width           =   888
   End
   Begin VB.TextBox txtTime 
      Height          =   372
      Left            =   2880
      TabIndex        =   8
      Top             =   1560
      Width           =   1452
   End
   Begin VB.TextBox txtPrimes 
      Height          =   372
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   1452
   End
   Begin VB.TextBox txtMaxPrime 
      Height          =   372
      Left            =   2880
      TabIndex        =   3
      Text            =   " 5000"
      Top             =   600
      Width           =   1452
   End
   Begin VB.TextBox txtIterate 
      Height          =   372
      Left            =   2880
      TabIndex        =   2
      Text            =   "5"
      Top             =   120
      Width           =   1452
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSieve 
      Caption         =   "&Sieve"
      Default         =   -1  'True
      Height          =   495
      Left            =   108
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Time (ms):"
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Primes:"
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Maximum Prime:"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Iterations:"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FSieveClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Change to true if you can get CSieveMFC registered on your system
#Const fMFC = 0

Enum ESieveType
    estBasicLocalFunction
    estBasicLocalClass
    estBasicGlobalFunctionP
    estBasicGlobalFunctionN
    estBasicDllPCode
    estBasicDllNative
    estBasicCtlPCode
    estBasicCtlNative
    estBasicExePCode
    estBasicExeNative
    estCppATL
    estCppMFC
End Enum

Private Declare Function timeGetTime Lib "winmm" () As Long

Private fDisplay As Boolean
Private dx As Long, dxOut As Long

Private Sub Form_Load()
    cboServer.Text = cboServer.List(0)
    dxOut = lstOutput.Left + Width - ScaleWidth
    dx = Width
    Width = dxOut
#If fMFC Then
    cboServer.AddItem "C++ MFC DLL"
#End If
End Sub

Private Sub cboServer_Click()
    Select Case cboServer.ListIndex
    Case estBasicLocalFunction
        chkAll.Enabled = False
        chkAll.Value = vbUnchecked
        chkLate.Enabled = False
    Case estBasicCtlPCode, estBasicCtlNative
        chkLate.Enabled = False
        chkLate.Value = vbUnchecked
        chkAll.Enabled = True
    Case estBasicGlobalFunctionP, estBasicGlobalFunctionN
        chkLate.Enabled = False
        chkLate.Value = vbUnchecked
        chkAll.Enabled = False
        chkLate.Value = vbUnchecked
    Case Else
        chkAll.Enabled = True
        chkLate.Enabled = True
    End Select
End Sub

Private Sub chkDisplay_Click()
    Static cLastIter As Integer
    If cLastIter = 0 Then cLastIter = txtIterate.Text
    fDisplay = (chkDisplay.Value = vbChecked)
    If fDisplay Then
        txtIterate.Text = 1
        Width = dx
    Else
        txtIterate.Text = cLastIter
        Width = dxOut
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSieve_Click()
    Dim ms As Long, i As Integer, iPrime As Integer, cPrime As Integer
    Dim ai() As Integer, vai As Variant
       
    ' Initialize prime variables
    Dim iMaxPrime As Integer, cIter As Integer, cPrimeMax As Integer
    iMaxPrime = txtMaxPrime.Text
    cIter = txtIterate.Text
    cPrimeMax = txtMaxPrime.Text
    txtTime.Text = ""
    txtPrimes.Text = ""
    txtTime.Refresh
    txtPrimes.Refresh
    
    Dim mpcMouse As MousePointerConstants
    mpcMouse = MousePointer
    MousePointer = vbHourglass
    
    ' Default early binding
    If chkLate = vbUnchecked Then
        Select Case cboServer.ListIndex
        Case estBasicLocalFunction
            ' Get all at once
            ms = timeGetTime()
            For i = 1 To cIter
                ReDim ai(0 To cPrimeMax)
                cPrime = Sieve(ai())
                If fDisplay Then
                    lstOutput.Clear
                    For iPrime = 0 To cPrime - 1
                        lstOutput.AddItem ai(iPrime)
                        lstOutput.TopIndex = lstOutput.ListCount - 1
                        lstOutput.Refresh
                    Next
                End If
            Next
            txtTime.Text = timeGetTime() - ms
            txtPrimes.Text = cPrime
        Case estBasicLocalClass
            ' Basic local class
            Dim sieveLocal As New CSieve
            sieveLocal.MaxPrime = txtMaxPrime.Text
            If chkAll = vbUnchecked Then
                ' Get one at a time
                ms = timeGetTime()
                For i = 1 To cIter
                    sieveLocal.ReInitialize
                    Do
                        iPrime = sieveLocal.NextPrime
                        If fDisplay And iPrime Then
                            lstOutput.AddItem iPrime
                            lstOutput.TopIndex = lstOutput.ListCount - 1
                            lstOutput.Refresh
                        End If
                    Loop Until iPrime = 0
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveLocal.Primes
            Else
                ' Get all at once
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    sieveLocal.AllPrimes ai()
                    If fDisplay Then
                        For iPrime = 0 To sieveLocal.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveLocal.Primes
            End If
        Case estBasicGlobalFunctionP
            ' Global function
            ms = timeGetTime()
            For i = 1 To cIter
                ReDim ai(0 To cPrimeMax)
                cPrime = SieveGlobalP(ai())
                If fDisplay Then
                    lstOutput.Clear
                    For iPrime = 0 To cPrime - 1
                        lstOutput.AddItem ai(iPrime)
                    Next
                End If
            Next
            txtTime.Text = timeGetTime() - ms
            txtPrimes.Text = cPrime
        Case estBasicGlobalFunctionN
            ' Global function
            ms = timeGetTime()
            For i = 1 To cIter
                ReDim ai(0 To cPrimeMax)
                cPrime = SieveGlobalN(ai())
                If fDisplay Then
                    lstOutput.Clear
                    For iPrime = 0 To cPrime - 1
                        lstOutput.AddItem ai(iPrime)
                    Next
                End If
            Next
            txtTime.Text = timeGetTime() - ms
            txtPrimes.Text = cPrime
        Case estBasicDllPCode
            ' Basic p-code DLL version, early bind
            Dim SieveBasDllP As New CSieveBasDllP
            SieveBasDllP.MaxPrime = txtMaxPrime.Text
            If chkAll = vbUnchecked Then
                ' Get one at a time
                ms = timeGetTime()
                For i = 1 To cIter
                    SieveBasDllP.ReInitialize
                    Do
                        iPrime = SieveBasDllP.NextPrime
                        If fDisplay And iPrime Then
                            lstOutput.AddItem iPrime
                            lstOutput.TopIndex = lstOutput.ListCount - 1
                            lstOutput.Refresh
                        End If
                    Loop Until iPrime = 0
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = SieveBasDllP.Primes
            Else
                ' Get all at once
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    SieveBasDllP.AllPrimes ai()
                    If fDisplay Then
                        For iPrime = 0 To SieveBasDllP.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = SieveBasDllP.Primes
            End If
        Case estBasicDllNative
            ' Basic DLL version, early bind
            Dim SieveBasDllN As New CSieveBasDllN
            SieveBasDllN.MaxPrime = txtMaxPrime.Text
            If chkAll = vbUnchecked Then
                ' Get one at a time
                ms = timeGetTime()
                For i = 1 To cIter
                    SieveBasDllN.ReInitialize
                    Do
                        iPrime = SieveBasDllN.NextPrime
                        If fDisplay And iPrime Then
                            lstOutput.AddItem iPrime
                            lstOutput.TopIndex = lstOutput.ListCount - 1
                            lstOutput.Refresh
                        End If
                    Loop Until iPrime = 0
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = SieveBasDllN.Primes
            Else
                ' Get all at once
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    SieveBasDllN.AllPrimes ai()
                    If fDisplay Then
                        For iPrime = 0 To SieveBasDllN.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = SieveBasDllN.Primes
            End If
        Case estBasicCtlPCode
            ' Basic p-code control version
            sieveCtlP.MaxPrime = txtMaxPrime.Text
            If chkAll = vbUnchecked Then
                ' Get one at a time
                ms = timeGetTime()
                For i = 1 To cIter
                    sieveCtlP.ReInitialize
                    Do
                        iPrime = sieveCtlP.NextPrime
                        If fDisplay And iPrime Then
                            lstOutput.AddItem iPrime
                            lstOutput.TopIndex = lstOutput.ListCount - 1
                            lstOutput.Refresh
                        End If
                    Loop Until iPrime = 0
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveCtlP.Primes
            Else
                ' Get all at once
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    sieveCtlP.AllPrimes ai()
                    If fDisplay Then
                        For iPrime = 0 To sieveCtlP.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveCtlP.Primes
            End If
        Case estBasicCtlNative
            ' Basic native control version
            sieveCtlN.MaxPrime = txtMaxPrime.Text
            If chkAll = vbUnchecked Then
                ' Get one at a time
                ms = timeGetTime()
                For i = 1 To cIter
                    sieveCtlN.ReInitialize
                    Do
                        iPrime = sieveCtlN.NextPrime
                        If fDisplay And iPrime Then
                            lstOutput.AddItem iPrime
                            lstOutput.TopIndex = lstOutput.ListCount - 1
                            lstOutput.Refresh
                        End If
                    Loop Until iPrime = 0
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveCtlN.Primes
            Else
                ' Get all at once
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    sieveCtlN.AllPrimes ai()
                    If fDisplay Then
                        For iPrime = 0 To sieveCtlN.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveCtlN.Primes
            End If
        Case estBasicExePCode
            ' Basic p-code EXE version, early bind
            Dim sieveBasExeP As New CSieveBasExeP
            sieveBasExeP.MaxPrime = txtMaxPrime.Text
            If chkAll = vbUnchecked Then
                ' Get one at a time
                ms = timeGetTime()
                For i = 1 To cIter
                    sieveBasExeP.ReInitialize
                    Do
                        iPrime = sieveBasExeP.NextPrime
                        If fDisplay And iPrime Then
                            lstOutput.AddItem iPrime
                            lstOutput.TopIndex = lstOutput.ListCount - 1
                            lstOutput.Refresh
                        End If
                    Loop Until iPrime = 0
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveBasExeP.Primes
            Else
                ' Get all at once
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    sieveBasExeP.AllPrimes ai()
                    If fDisplay Then
                        For iPrime = 0 To sieveBasExeP.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveBasExeP.Primes
            End If
        Case estBasicExeNative
            ' Basic native EXE version, early bind
            Dim sieveBasExeN As New CSieveBasExeN
            sieveBasExeN.MaxPrime = txtMaxPrime.Text
            If chkAll = vbUnchecked Then
                ' Get one at a time
                ms = timeGetTime()
                For i = 1 To cIter
                    sieveBasExeN.ReInitialize
                    Do
                        iPrime = sieveBasExeN.NextPrime
                        If fDisplay And iPrime Then
                            lstOutput.AddItem iPrime
                            lstOutput.TopIndex = lstOutput.ListCount - 1
                            lstOutput.Refresh
                        End If
                    Loop Until iPrime = 0
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveBasExeN.Primes
            Else
                ' Get all at once
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    sieveBasExeN.AllPrimes ai()
                    If fDisplay Then
                        For iPrime = 0 To sieveBasExeN.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveBasExeN.Primes
            End If
        Case estCppATL
            Dim sieveAtl As CSieveATL
            Set sieveAtl = New CSieveATL
            sieveAtl.MaxPrime = txtMaxPrime.Text
            If chkAll = vbUnchecked Then
                ' Get one at a time
                ms = timeGetTime()
                For i = 1 To cIter
                    sieveAtl.ReInitialize
                    Do
                        iPrime = sieveAtl.NextPrime
                        If fDisplay And iPrime Then
                            lstOutput.AddItem iPrime
                            lstOutput.TopIndex = lstOutput.ListCount - 1
                            lstOutput.Refresh
                        End If
                    Loop Until iPrime = 0
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveAtl.Primes
            Else
                ' Get all at once
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    sieveAtl.AllPrimes ai()
                    If fDisplay Then
                        For iPrime = 0 To sieveAtl.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = sieveAtl.Primes
            End If
' MFC server registration is so flaky that I have commented this out.
' Put it back in if you can get the server registered on your system.
#If fMFC Then
        Case estCppMFC
            Dim SieveMFC As New CSieveMFC
            SieveMFC.MaxPrime = txtMaxPrime.Text
            If chkAll = vbUnchecked Then
                ' Get one at a time
                ms = timeGetTime()
                For i = 1 To cIter
                    SieveMFC.ReInitialize
                    Do
                        iPrime = SieveMFC.NextPrime
                        If fDisplay And iPrime Then
                            lstOutput.AddItem iPrime
                            lstOutput.TopIndex = lstOutput.ListCount - 1
                            lstOutput.Refresh
                        End If
                    Loop Until iPrime = 0
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = SieveMFC.Primes
            Else
                ' Get all at once
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    ' Put array in Variant for MFC
                    vai = ai()
                    SieveMFC.AllPrimes vai
                    If fDisplay Then
                        For iPrime = 0 To SieveMFC.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
                txtPrimes.Text = SieveMFC.Primes
            End If
#End If
        End Select
    Else ' Late bound
        ' Set variable at run time
        Dim sieveLate As Object
        Select Case cboServer.ListIndex
        Case estBasicLocalClass
            Set sieveLate = New CSieve
#Const fUseTypeLib = 1
#If fUseTypeLib Then
        Case estBasicDllPCode
            Set sieveLate = New CSieveBasDllP
        Case estBasicDllNative
            Set sieveLate = New CSieveBasDllN
        Case estBasicExePCode
            Set sieveLate = New CSieveBasExeP
        Case estBasicExeNative
            Set sieveLate = New CSieveBasExeN
        Case estCppATL
            Set sieveLate = New CSieveATL
#If fMFC Then
        Case estCppMFC
            Set sieveLate = New CSieveMFC
#End If
#Else
        Case estBasicEXE
            Set sieveLate = CreateObject("SieveBasDllP.CSieveBasDllP")
        Case estBasicDllNative
            Set sieveLate = CreateObject("SieveBasDllN.CSieveBasDllN")
        Case estBasicExePCode
            Set sieveLate = CreateObject("SieveBasExeP.CSieveBasExeP")
        Case estBasicExeNative
            Set sieveLate = CreateObject("SieveBasExeN.CSieveBasExeN")
        Case estCppATL
            Set sieveLate = CreateObject("SieveAtl.CSieveATL")
#If fMFC Then
        Case estCppMFC
            Set sieveLate = CreateObject("SieveMFC.CSieveMFC")
#End If
#End If
        End Select
        sieveLate.MaxPrime = txtMaxPrime.Text
        If chkAll = vbUnchecked Then
            ' Get one at a time
            ms = timeGetTime()
            For i = 1 To cIter
                sieveLate.ReInitialize
                Do
                    iPrime = sieveLate.NextPrime
                    If fDisplay And iPrime Then
                        lstOutput.AddItem iPrime
                        lstOutput.TopIndex = lstOutput.ListCount - 1
                        lstOutput.Refresh
                    End If
                Loop Until iPrime = 0
            Next
            txtTime.Text = timeGetTime() - ms
            txtPrimes.Text = sieveLate.Primes
        Else
            ' Get all at once
            If cboServer.ListIndex <> estCppMFC Then
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    sieveLate.AllPrimes ai()
                    If fDisplay Then
                        For iPrime = 0 To sieveLate.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
            Else
                ' MFC different because it can't handle Basic arrays
                ms = timeGetTime()
                For i = 1 To cIter
                    ReDim ai(0 To cPrimeMax)
                    vai = ai()
                    ' Put array in Variant for MFC
                    sieveLate.AllPrimes vai
                    If fDisplay Then
                        For iPrime = 0 To sieveLate.Primes - 1
                            lstOutput.AddItem ai(iPrime)
                        Next
                    End If
                Next
                txtTime.Text = timeGetTime() - ms
            End If
            txtPrimes.Text = sieveLate.Primes
        End If
    End If
    MousePointer = mpcMouse
End Sub

Sub RefreshControls()
    fDisplay = (chkDisplay.Value = vbChecked)
    If lstOutput.ListCount Then lstOutput.Clear
End Sub


