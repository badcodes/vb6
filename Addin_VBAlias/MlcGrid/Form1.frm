VERSION 5.00
Object = "*\AProyecto1.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdDelRows 
      Caption         =   "DelRows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9840
      TabIndex        =   40
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton CmdDelGrid 
      Caption         =   "DelGrid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9840
      TabIndex        =   39
      Top             =   5640
      Width           =   1935
   End
   Begin MlcFlexGrid.MlcGrid MlcGrid1 
      Height          =   4815
      Left            =   2160
      TabIndex        =   38
      Top             =   1200
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Cols            =   2
      Row             =   0
      Col             =   0
   End
   Begin VB.CommandButton CmdChangeColPos 
      Caption         =   "ChangeColPos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7800
      TabIndex        =   37
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton CmdClearCol 
      Caption         =   "ClearCol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9840
      TabIndex        =   36
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton CmdClearRow 
      Caption         =   "ClearRow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9840
      TabIndex        =   35
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton CmdSorted 
      Caption         =   "Sorted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7800
      TabIndex        =   34
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton CmdAllowUserSortCol 
      Caption         =   "AllowUserSortCol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      TabIndex        =   33
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton CmdAllowUserChangeColPos 
      Caption         =   "AllowUserChangeColPos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      TabIndex        =   32
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton CmdAutoScrolls 
      Caption         =   "AutoScrolls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   31
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmdSelRow 
      Caption         =   "SelRow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   30
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton CmdFocusRect 
      Caption         =   "FocusRect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton CmdScrollTrack 
      Caption         =   "ScrollTrack"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   28
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CmdScroll 
      Caption         =   "Scrolls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   27
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdColWidth 
      Caption         =   "ColWidth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8160
      TabIndex        =   26
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdTextMatrix 
      Caption         =   "TextMatrix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   25
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton CmdText 
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   24
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton CmdRowHeight 
      Caption         =   "RowHeight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8160
      TabIndex        =   23
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CmdTitColor 
      Caption         =   "TituloBackColor        TituloForeColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   22
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton CmdSel 
      Caption         =   "SelBackColor          SelForecolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   21
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9840
      TabIndex        =   20
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton CmdRemove 
      Caption         =   "RemoveItem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9840
      TabIndex        =   19
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton CmdResetProperties 
      Caption         =   "ResetProperties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9840
      TabIndex        =   18
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton CmdShowGrid 
      Caption         =   "ShowGrid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton CmdGridColor 
      Caption         =   "GridColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton CmdForeColor 
      Caption         =   "Forecolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6360
      TabIndex        =   15
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton CmdBackcolor 
      Caption         =   "Backcolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6360
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdCellFontName 
      Caption         =   "CellFontName           CellFontsize              CellFontbold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton CmdCellForecolor 
      Caption         =   "CellForecolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton CmdCellBackcolor 
      Caption         =   "CellBackcolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton CmdCellsBackcolor 
      Caption         =   "CellsBackcolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "2 - AddItem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton CmdTitulo 
      Caption         =   "1 - Titulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label LMouseCol 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3240
      TabIndex        =   9
      Top             =   480
      Width           =   750
   End
   Begin VB.Label LMouseRow 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   750
   End
   Begin VB.Label LCol 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   5355
      TabIndex        =   7
      Top             =   420
      Width           =   750
   End
   Begin VB.Label LRow 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   4515
      TabIndex        =   6
      Top             =   420
      Width           =   750
   End
   Begin VB.Label Label4 
      Caption         =   "Col"
      Height          =   225
      Left            =   5565
      TabIndex        =   5
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label3 
      Caption         =   "Row"
      Height          =   225
      Left            =   4725
      TabIndex        =   4
      Top             =   210
      Width           =   540
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "MouseCol"
      Height          =   255
      Left            =   3255
      TabIndex        =   3
      Top             =   210
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MouseRow"
      Height          =   255
      Left            =   2310
      TabIndex        =   2
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAdd_Click()
    For i = 1 To 20
        Randomize
        mivalor = Int((1000000000000# * Rnd) + 1)
        MlcGrid1.AddItem mivalor & Chr(9) & mivalor / 2 & Chr(9) & i + 100 & vbTab & i + 200 & vbTab & i * 27 & vbTab & i * 42 & vbTab & i + 3 & vbTab & i + 2
    Next i
    MlcGrid1.Row = 1
    MlcGrid1.Col = 1
End Sub

Private Sub CmdAllowUserChangeColPos_Click()
    If MlcGrid1.AllowUserChangeColPos = True Then MlcGrid1.AllowUserChangeColPos = False Else MlcGrid1.AllowUserChangeColPos = True
End Sub

Private Sub CmdAllowUserSortCol_Click()
    If MlcGrid1.AllowUserSortCol = True Then MlcGrid1.AllowUserSortCol = False Else MlcGrid1.AllowUserSortCol = True
End Sub


Private Sub CmdAutoScrolls_Click()
    If MlcGrid1.AutoScrolls = False Then MlcGrid1.AutoScrolls = True Else MlcGrid1.AutoScrolls = False
End Sub

Private Sub CmdBackcolor_Click()
Dim Color As Single
Randomize
Color = Int((12 * Rnd) + 1)
    MlcGrid1.BackColor = QBColor(Color)
End Sub

Private Sub CmdCellBackcolor_Click()
    For i = 1 To 20
        Randomize
        mivalor = Int((20 * Rnd) + 1)
        MlcGrid1.CellBackColor(mivalor, 1) = vbGreen
        MlcGrid1.CellBackColor(mivalor, 2) = vbCyan
    Next i
End Sub

Private Sub CmdCellFontName_Click()
    For i = 1 To 20
        Randomize
        mivalor = Int((20 * Rnd) + 1)
        MlcGrid1.CellFontName(mivalor, 5) = "Comic sans ms"
        MlcGrid1.CellFontSize(mivalor, 5) = 12
        MlcGrid1.CellFontBold(mivalor, 5) = True
    Next i
End Sub

Private Sub CmdCellForecolor_Click()
    For i = 1 To 20
        Randomize
        mivalor = Int((20 * Rnd) + 1)
        MlcGrid1.CellForeColor(mivalor, 3) = vbRed
        MlcGrid1.CellForeColor(mivalor, 4) = vbBlue
    Next i
End Sub

Private Sub CmdCellsBackcolor_Click()
Dim Color As Single
Randomize
Color = Int((12 * Rnd))
    MlcGrid1.CellsBackColor = QBColor(Color)
End Sub

Private Sub CmdChangeColPos_Click()
    MlcGrid1.ChangeColPosition 3, 1
End Sub

Private Sub CmdClear_Click()
    MlcGrid1.Clear
End Sub

Private Sub CmdClearCol_Click()
    MlcGrid1.ClearCol 3, False
End Sub

Private Sub CmdClearRow_Click()
    MlcGrid1.ClearRow 5, False
End Sub

Private Sub CmdColWidth_Click()
Dim Ancho As Integer
    For i = 1 To MlcGrid1.Cols
        Randomize
        Ancho = Int((2000 * Rnd) + 1)
        MlcGrid1.ColWidth(i) = Ancho
    Next i
End Sub

Private Sub CmdDelGrid_Click()
    MlcGrid1.DelGrid
End Sub

Private Sub CmdDelRows_Click()
    MlcGrid1.DelRows
End Sub

Private Sub CmdFocusRect_Click()
Static Focus As Integer
    If Focus < 2 Then Focus = Focus + 1 Else Focus = 0
    MlcGrid1.FocusRect = Focus
End Sub

Private Sub CmdForeColor_Click()
Dim Color As Single
Randomize
Color = Int((12 * Rnd) + 1)
    MlcGrid1.ForeColor = QBColor(Color)
End Sub

Private Sub CmdGridColor_Click()
Dim Color As Single
Randomize
Color = Int((12 * Rnd) + 1)
    MlcGrid1.GridColor = QBColor(Color)
End Sub

Private Sub CmdRemove_Click()
    If MlcGrid1.Rows >= 3 Then
        MlcGrid1.RemoveItem 3
    Else
        MsgBox "Se intenta eliminar una fila no existente.", vbInformation, "ERROR"
    End If
End Sub

Private Sub CmdResetProperties_Click()
    MlcGrid1.ResetProperties
End Sub

Private Sub CmdRowHeight_Click()
Dim Alto As Integer
Randomize
Alto = Int((990 * Rnd) + 1)
    MlcGrid1.RowHeight = Alto
End Sub

Private Sub CmdScroll_Click()
Static Scroll As Integer
    If Scroll < 3 Then Scroll = Scroll + 1 Else Scroll = 0
    MlcGrid1.Scrolls = Scroll
End Sub

Private Sub CmdScrollTrack_Click()
    If MlcGrid1.ScrollTrack = False Then MlcGrid1.ScrollTrack = True Else MlcGrid1.ScrollTrack = False
End Sub

Private Sub CmdSel_Click()
    MlcGrid1.SelBackColor = &H80FF&
    MlcGrid1.SelForeColor = vbYellow
End Sub

Private Sub CmdSelRow_Click()
    If MlcGrid1.SelRow = True Then MlcGrid1.SelRow = False Else MlcGrid1.SelRow = True
End Sub

Private Sub CmdShowGrid_Click()
Static Grid As Boolean
    If Grid = True Then Grid = False Else Grid = True
    MlcGrid1.ShowGrid = Grid
End Sub

Private Sub CmdSorted_Click()
Static SortCol As Boolean
    If SortCol = True Then SortCol = False Else SortCol = True
    MlcGrid1.Sorted 2, 6, 1, SortCol
End Sub

Private Sub CmdText_Click()
    MsgBox "Texto de la Fila(Row) " & MlcGrid1.Row & " y de la Columna(Col) " & MlcGrid1.Col & " : " & MlcGrid1.Text, vbInformation, "Propiedad Text"
End Sub

Private Sub CmdTextMatrix_Click()
    MsgBox "Texto de la Fila(Row) 3 y de la Columna(Col) 3 : " & MlcGrid1.TextMatrix(3, 3), vbInformation, "Propiedad TextMatrix"
End Sub

Private Sub CmdTitColor_Click()
    MlcGrid1.TituloBackColor = vbRed
    MlcGrid1.TituloForeColor = vbWhite
End Sub

Private Sub CmdTitulo_Click()
    MlcGrid1.Titulo = "<Codigo de Barras                |<Referencia     |^Familia         |^Proveedor    |>Coste       |>P.V.P.       |>Stock     |>Reservado"
End Sub


Private Sub MlcGrid1_EnterCell()
    LRow.Caption = MlcGrid1.Row
    LRow.Refresh
    LCol.Caption = MlcGrid1.Col
    LCol.Refresh
End Sub

Private Sub MlcGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LMouseRow.Caption = MlcGrid1.MouseRow
    LMouseCol.Caption = MlcGrid1.MouseCol
End Sub

