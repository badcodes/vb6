VERSION 5.00
Begin VB.Form FSieveBasExeP 
   Caption         =   "Sieve of Eratosthenes P-Code Server"
   ClientHeight    =   2412
   ClientLeft      =   3408
   ClientTop       =   1740
   ClientWidth     =   5628
   Icon            =   "SieveBasExeP.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2412
   ScaleWidth      =   5628
   Begin VB.CheckBox chkAll 
      Caption         =   "Get All"
      Height          =   255
      Left            =   3540
      TabIndex        =   13
      Top             =   2040
      Width           =   828
   End
   Begin VB.CheckBox chkLate 
      Caption         =   "Late Bind"
      Height          =   255
      Left            =   2496
      TabIndex        =   12
      Top             =   2040
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
      Left            =   1536
      TabIndex        =   10
      Top             =   2040
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Time:"
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
Attribute VB_Name = "FSieveBasExeP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm" () As Long

Private fDisplay As Boolean
Private dx As Long, dxOut As Long

Private Sub Form_Load()
    dxOut = lstOutput.Left + Width - ScaleWidth
    dx = Width
    Width = dxOut
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
        ' Basic native EXE version, early bind
        Dim sieveBasExeP As New CSieveBasExeP
        sieveBasExeP.MaxPrime = txtMaxPrime.Text
        ' Get one at a time
        If chkAll = vbUnchecked Then
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
                ReDim ai(0 To 0)    ' Zero array
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
    Else ' Late bound
        ' Set variable at run time
        Dim sieveLate As Object
        Set sieveLate = New CSieveBasExeP
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
            ms = timeGetTime()
            For i = 1 To cIter
                ReDim ai(0 To 0)    ' Zero array
                ReDim ai(0 To cPrimeMax)
                sieveLate.AllPrimes ai()
                If fDisplay Then
                    For iPrime = 0 To sieveLate.Primes - 1
                        lstOutput.AddItem ai(iPrime)
                    Next
                End If
            Next
            txtTime.Text = timeGetTime() - ms
            txtPrimes.Text = sieveLate.Primes
        End If
    End If
    MousePointer = mpcMouse
End Sub

Sub RefreshControls()
    fDisplay = (chkDisplay.Value = vbChecked)
    If lstOutput.ListCount Then lstOutput.Clear
End Sub

