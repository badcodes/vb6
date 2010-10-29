VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "MSWLESS.OCX"
Begin VB.Form frmPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   5235
   ClientTop       =   4350
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1215
      Left            =   540
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin MSWLess.WLFrame WLFrame1 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5636
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mTP As cTablePrint
Attribute mTP.VB_VarHelpID = -1


Public Sub PrintMacroGrid()
Dim fntSet As IFont, fntSet2 As StdFont
Dim i As Long, j As Long

Me.Cls

With mTP
    .HasFooter = False

'Set how far the text should keep away from the cells' borders
    .CellXOffset = 60
    .CellYOffset = 15
    .HeaderRows = 1
    .HeaderRowHeightMin = 600
    .Cols = 5
    .Rows = frmAddIn.MlcGrid1.Rows / 2 + 1

'Set the fntSet to the form's Font
    Set fntSet = Me.Font
    Set .FontMatrix(-1, -1) = Me.Font

'Clone fntSet (that's why I'm using IFont (fully compatible to StdFont) here ;-)
    fntSet.Clone fntSet2

'Make the font bold and bigger
    fntSet2.Bold = True
    fntSet2.Size = 10

'Set the font for the page header
    Set .HeaderFont(-1, -1) = fntSet2
    Set .FooterFont(-1) = fntSet2 'not used


    .HeaderLineThickness = 2
    .ColAlignment(0) = eCenter 'eLeft
    .ColAlignment(1) = eCenter
    .ColAlignment(3) = eCenter 'eLeft
    .ColAlignment(4) = eCenter
    .MarginTop = 450
    .MarginLeft = 960
    .PrintHeaderOnEveryPage = True

    .HeaderText(-1, 0) = "The string"
    .HeaderText(-1, 1) = "Will be replaced with"
    .HeaderText(-1, 3) = "The string"
    .HeaderText(-1, 4) = "Will be replaced with"


    For i = 1 To frmAddIn.MlcGrid1.Rows Step 2
        If i > frmAddIn.MlcGrid1.Rows Then Exit For

        frmAddIn.MlcGrid1.Row = i

        frmAddIn.MlcGrid1.Col = 1
        .TextMatrix((i - 1) / 2, 0) = frmAddIn.MlcGrid1.Text

        frmAddIn.MlcGrid1.Col = 2
        .TextMatrix((i - 1) / 2, 1) = frmAddIn.MlcGrid1.Text

        If i + 1 > frmAddIn.MlcGrid1.Rows Then Exit For
        frmAddIn.MlcGrid1.Row = i + 1

        frmAddIn.MlcGrid1.Col = 1
        .TextMatrix((i - 1) / 2, 3) = frmAddIn.MlcGrid1.Text

        frmAddIn.MlcGrid1.Col = 2
        .TextMatrix((i - 1) / 2, 4) = frmAddIn.MlcGrid1.Text

    Next


    .ColWidth(0) = 1200
    .ColWidth(1) = 3600
    .ColWidth(2) = 60 '(Me.ScaleWidth - 300) / .Cols
    .ColWidth(3) = 1200
    .ColWidth(4) = 3600

'Finally draw the Grid on the form:
' (Simply change "Me" to "Printer" to print it on the printer !)
    .DrawTable Printer
'.DrawTable Me

'Print a reference and page numbers
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print Space$(23) & "Vb IDE Extender" & Space$(159) & " MACROS, Page " & Printer.Page
    Unload Me

'To see on the form
'Print
'Print
'Print
'Print
'Print
'Print
'Print Space$(23) & "Vb IDE Extender" & Space$(158) & " MACROS, Page 1"

End With
End Sub


Private Sub Form_Load()

Set mTP = New cTablePrint

With Timer1
    .Interval = 1000
    .Enabled = True
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mTP = Nothing

End Sub


Private Sub mTP_NewPage(objOutput As Object, TopMarginAlreadySet As Boolean, bCancel As Boolean, ByVal lLastPrintedRow As Long)
'The form is simply cleared in this example here.
'You should not do this in your programs !
'If you print the table on a printer, simply call Printer.NewPage in here.
'If you're making a kind of page preview, you will have to create some
'sort of multi-page mechanism. (For example, caching all pages as bitmaps (simple but slow) or
'setting bCancel = True and using lLastPrintedRow + 1 as the lRowToStart parameter to DrawTable()
'when drawing the next page, etc.)

'Set TopMarginAlreadySet = True if objOutput.CurrentY is the position where
'the next part of the grid should start. Otherwise the value from MarginTop
'is added to objOutput.CurrentY.

'Me.Cls


Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Space$(23) & "Vb IDE Extender" & Space$(159) & " MACROS, Page " & Printer.Page
Printer.NewPage
End Sub


Private Sub Timer1_Timer()

Timer1.Enabled = False
PrintMacroGrid
DoEvents
Unload frmPrint

End Sub

