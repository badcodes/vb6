VERSION 5.00
Begin VB.UserControl MlcGrid 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   Picture         =   "MlcGrid.ctx":0000
   ScaleHeight     =   1560
   ScaleWidth      =   1890
   ToolboxBitmap   =   "MlcGrid.ctx":08CA
   Begin VB.PictureBox PFlex 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1485
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   1740
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      Begin VB.CommandButton CLeft 
         Height          =   225
         Left            =   1080
         Picture         =   "MlcGrid.ctx":0BDC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   225
      End
      Begin VB.CommandButton CTop 
         Height          =   225
         Left            =   720
         Picture         =   "MlcGrid.ctx":0F66
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   225
      End
      Begin VB.HScrollBar HsHorizontal 
         Height          =   225
         LargeChange     =   10
         Left            =   0
         TabIndex        =   2
         Top             =   1155
         Value           =   1
         Width           =   1275
      End
      Begin VB.VScrollBar VsVertical 
         Height          =   1065
         LargeChange     =   10
         Left            =   1470
         TabIndex        =   1
         Top             =   0
         Width           =   225
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1470
         Top             =   1155
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "MlcGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'''''''''''''''''''''''''
'   Miguel Lirón        '
'                       '
'   lairon@menta.net    '
'''''''''''''''''''''''''

Public Enum TipoScroll
    Ninguna
    Horizontal
    Vertical
    Ambas
End Enum
Private ValorScroll As TipoScroll

Public Enum TipoFocus
    Light
    Heavy
    None
End Enum
Private ValorFocus As TipoFocus

Private LRows As Integer
Private LCols As Integer

Private LRow As Integer
Private LCol As Integer

Private VirtualHeight As Integer
Private VirtualTitulo As String
Private VirtualMouseRow As Integer
Private VirtualMouseCol As Integer
Private LShowGrid As Boolean 'determina se si verá el grid o no.
Private LBackcolor As OLE_COLOR 'Backcolor
Private LCellsBackColor As OLE_COLOR 'Color de las celdas
Private LGridColor As OLE_COLOR 'Color del grid.
Private LForeColor As OLE_COLOR 'Color del texto de las celdas.
Private LTituloForeColor As OLE_COLOR 'Color del texto del titulo.
Private LTituloBackColor As OLE_COLOR 'Color de las celdas del titulo.
Private LSelBackColor As OLE_COLOR 'Color de fondo de las celdas seleccionadas.
Private LSelForeColor As OLE_COLOR 'Color del texto seleccionado.
Private LSeeAllText As Boolean 'Determina si se vera todo el texto de la celda o no.
Private LScrollTrack As Boolean 'Determina si se desplazan las celdas al desplazar las barras de scroll.
Private LSelRow As Boolean 'Determina si se seleccionara toda la fila o no.
Private LAutoScrolls As Boolean 'Determina si se haran visibles los scrolls cuando el ancho o el alto sea superior a las medidas del flex.
Private LAllowUserChangeColPos As Boolean 'Determina si el usuario puede o no cambiar las columnas de posicion.
Private LAllowUserSortCol As Boolean 'Determina si el usuario puede o no ordenar las columnas.

Dim ColeccionTitulo As New Collection 'Donde guardo el Titulo
Dim ColeccionColWidth As New Collection 'Donde guardo el width de cada columna.
Dim ColeccionAlineacion As New Collection 'Donde guardo la alineación de las columnas.
Dim ArrayRows As Variant 'donde guardo los datos de las filas uno elemento por fila.
Dim ArrayCols As Variant 'Donde guardo los datos de cada fila un elemento por columna.
Dim ArrayCell As Variant 'Donde guardo las propiedades de cada celda.

Dim AntCol As Integer
Dim AntRow As Integer

Dim RowsCompletas As Boolean 'Indica si se han dibujado o no todas las filas.
Dim Adding As Boolean 'Indica si estoy añadiendo filas o no.
Dim AnchoTotal As Integer 'Ancho para dibujar el cuadro blanco.
Dim AltoTotal As Integer 'Alto para dibujar el cuadro blanco.

Dim UltimaFila As Integer 'Ultima fila dibujada.
Dim UltimaColumna As Integer 'Ultima columna dibujada.
Dim ClickRowCol As Boolean 'Indica si cambiamos la celda haciendo click o utilizando otro metodo.

Public Event Click()
Public Event DblClick()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event EnterCell()
Public Event LeaveCell()

Const LColWidth = 960 'Ancho mínimo de las columnas.
'Dibuja la celda.
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'dibuja el enfoque de la celda
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
'Dibuja el texto en el rectangulo definido.
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
        (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
        lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
'Private Const DT_TOP = &H00000000
'Private Const DT_LEFT = &H00000000
'Private Const DT_CENTER = &H00000001
'Private Const DT_RIGHT = &H00000002
'Private Const DT_VCENTER = &H00000004
'Private Const DT_BOTTOM = &H00000008
'Private Const DT_WORDBREAK = &H00000010
'Private Const DT_SINGLELINE = &H00000020
'Private Const DT_EXPANDTABS = &H00000040
'Private Const DT_TABSTOP = &H00000080
'Private Const DT_NOCLIP = &H00000100
'Private Const DT_EXTERNALLEADING = &H00000200
'Private Const DT_CALCRECT = &H00000400
'Private Const DT_NOPREFIX = &H00000800
'Private Const DT_INTERNAL = &H00001000
'Private Const DT_EDITCONTROL = &H00002000
'Private Const DT_PATH_ELLIPSIS = &H00004000
'Private Const DT_END_ELLIPSIS = &H00008000
'Private Const DT_MODIFYSTRING = &H00010000
'Private Const DT_RTLREADING = &H00020000
'Private Const DT_WORD_ELLIPSIS = &H00040000

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Function AddItem(item As Variant, Optional ByVal Index As Integer)
Dim i As Integer, Inicio As Integer, Mipos As Integer, TotalVbTab As Integer
Dim ArrayCols1 As Variant, FilaACambiar As Integer
Inicio = 1
Mipos = 1
Adding = True

'Si item no tiene valor.
If Trim(item) = "" Then
    ReDim ArrayCols(LCols)
    For i = 1 To LCols
        ReDim ArrayCell(10)
        ArrayCols(i) = ArrayCell
    Next i
Else
'quito todos los vbtab finales de la cadena item.
    Do
        If Mid(item, Len(item), 1) <> vbTab Then Exit Do
        If Mid(item, Len(item), 1) = vbTab Then
            item = Mid(item, 1, Len(item) - 1)
        End If
    Loop

    Mipos = InStr(Inicio, item, vbTab)
    If Mipos = 0 Then 'si solo se llena una columna.
        ReDim ArrayCols(LCols)
        For i = 1 To LCols
            ReDim ArrayCell(10)
            If i = 1 Then ArrayCell(1) = item
            ArrayCols(i) = ArrayCell
        Next i
    Else
        ReDim ArrayCols(LCols)
        Do
            Mipos = InStr(Inicio, item, vbTab)
            TotalVbTab = TotalVbTab + 1
            If Mipos = 0 Then
                ReDim ArrayCell(10)
                ArrayCell(1) = Mid(item, Inicio, Len(item) - Inicio + 1)
                ArrayCols(TotalVbTab) = ArrayCell
                Exit Do
            End If
            If TotalVbTab = 1 Then
                ReDim ArrayCell(10)
                ArrayCell(1) = Mid(item, Inicio, Mipos - 1)
                ArrayCols(1) = ArrayCell
            Else
                ReDim ArrayCell(10)
                ArrayCell(1) = Mid(item, Inicio, Mipos - Inicio)
                ArrayCols(TotalVbTab) = ArrayCell
            End If
            Inicio = Mipos + 1
            If TotalVbTab = LCols Then
                Exit Do
            End If
        Loop
'lleno el resto de arraycols, es decir TotalVbTab ha sido menor que LCols.
        For i = TotalVbTab + 1 To LCols
            ReDim ArrayCell(10)
            ArrayCols(i) = ArrayCell
        Next i
    End If

End If
''''''''''''''''''''''''''''''''''''''''
Select Case Index
Case Is = 0 'No se especifica indice.
    If LRows = 0 Then
        ReDim ArrayRows(1)
        ArrayRows(1) = ArrayCols
    Else
        ReDim Preserve ArrayRows(LRows + 1)
        ArrayRows(UBound(ArrayRows)) = ArrayCols
    End If
Case Else 'Cualquier indice.
'Si está en la cuadricula, redibujamos la cuadricula.
    If Index >= VsVertical.Value And Index <= UltimaFila Then RowsCompletas = False
    ReDim Preserve ArrayRows(LRows + 1)
    For i = 1 To UBound(ArrayRows)
        FilaACambiar = UBound(ArrayRows) - i
        If FilaACambiar >= Index Then
            ArrayCols1 = ArrayRows(FilaACambiar)
            ArrayRows(FilaACambiar + 1) = ArrayCols1
        End If
    Next i
    ArrayRows(Index) = ArrayCols
End Select
''''''''''''''''''''''''''''''''''''''''
Rows = Rows + 1
End Function

Private Function CalcularAnchoTotal()
Dim AnchoParcial As Integer, Columna As Integer
AnchoTotal = 0
For Columna = HsHorizontal.Value To LCols
    If ColeccionColWidth.Count > 0 And Columna > 0 Then
        If ColeccionColWidth.item(Columna) = "" Then
            AnchoParcial = AnchoParcial + LColWidth
        Else
            AnchoParcial = AnchoParcial + ColeccionColWidth.item(Columna)
        End If
    Else
        AnchoParcial = AnchoParcial + LColWidth
    End If
    If AnchoParcial > PFlex.ScaleWidth Then
        AnchoTotal = PFlex.ScaleWidth
'Compruebo autoscrolls.
        If LAutoScrolls = True Then
            If Scrolls <> Ambas And Scrolls <> Horizontal Then
                If Scrolls = Vertical Then
                    Scrolls = Ambas
                Else
                    Scrolls = Horizontal
                End If
            End If
        End If
'fin comprobacion.
        Exit For
    Else
        AnchoTotal = AnchoParcial
    End If
Next Columna
End Function

Private Function DibujarCeldas()
Dim Fila As Integer, Columna As Integer, X As Long, Y As Long, AnchoColumna As Integer
Dim ValorHorizontal As Integer, ValorVertical As Integer, i As Integer
Dim lSuccess As Long, MyRect As RECT, Negrita As Boolean

ValorHorizontal = IIf(HsHorizontal.Value = 0, 1, HsHorizontal.Value)
ValorVertical = IIf(VsVertical.Value = 0, 1, VsVertical.Value)

Call DibujarTitulo
If Rows = 0 Then Exit Function
PFlex.ScaleMode = vbPixels

RowsCompletas = IIf(AltoTotal / 15 >= PFlex.ScaleHeight, True, False)

Y = VirtualHeight / 15
X = 0
For Fila = ValorVertical To LRows
    ArrayCols = ArrayRows(Fila)

    For Columna = ValorHorizontal To LCols
        PFlex.FontName = UserControl.FontName
        PFlex.Font.Size = UserControl.FontSize
        PFlex.Font.Bold = UserControl.FontBold
'Obtengo los datos de la celda.
        ArrayCell = ArrayCols(Columna)
'Ancho de la columna.
        AnchoColumna = IIf(ColeccionColWidth.Count > 0, ColeccionColWidth.item(Columna) / 15, LColWidth / 15)
'Fuente de la celda.
        If ArrayCell(4) <> "" Then PFlex.Font = ArrayCell(4)
'Tamaño de la fuente de la celda.
        If ArrayCell(5) <> "" Then PFlex.Font.Size = ArrayCell(5)
'Bold de la fuente
        Negrita = ArrayCell(6)
        PFlex.Font.Bold = ArrayCell(6)
'Color del marco de la celda.
        PFlex.ForeColor = IIf(ShowGrid = True, LGridColor, LCellsBackColor)
'Color de la celda.
        PFlex.FillColor = IIf(Fila = LRow And Columna <> LCol And LSelRow = True, SelBackColor, IIf(ArrayCell(2) <> "", ArrayCell(2), LCellsBackColor))
'Dibujo la celda.
        Rectangle PFlex.hDC, X, Y, X + 1 + AnchoColumna, Y + 1 + (VirtualHeight / 15)
'Escribo los datos.
        MyRect.Left = X + 3
        MyRect.Right = X + AnchoColumna - 3
        MyRect.Top = Y + 1
        MyRect.Bottom = Y + (VirtualHeight / 15)
'Color del texto dependiendo si esta la celda seleccionada o no.
        If Fila = LRow And Columna <> LCol And LSelRow = True Then
            PFlex.ForeColor = LSelForeColor
        Else
'Color del texto de la celda especificada.
            PFlex.ForeColor = IIf(ArrayCell(3) <> "", ArrayCell(3), LForeColor)
        End If
'lSuccess = DrawText(PFlex.hdc, ArrayCell(1), Len(ArrayCell(1)), MyRect, DT_VCENTER Or IIf(ColeccionAlineacion.Item(Columna) = "<", DT_LEFT, IIf(ColeccionAlineacion.Item(Columna) = ">", DT_RIGHT, DT_CENTER)) Or IIf(LSeeAllText = True, DT_WORDBREAK, DT_SINGLELINE))
        Select Case ColeccionAlineacion.item(Columna)
        Case "<" 'Alineacion a la izquierda.
            lSuccess = DrawText(PFlex.hDC, ArrayCell(1), Len(ArrayCell(1)), MyRect, DT_VCENTER Or DT_LEFT Or IIf(LSeeAllText = True, DT_WORDBREAK, DT_SINGLELINE))
        Case ">" 'alineacion a la derecha
            lSuccess = DrawText(PFlex.hDC, ArrayCell(1), Len(ArrayCell(1)), MyRect, DT_VCENTER Or DT_RIGHT Or IIf(LSeeAllText = True, DT_WORDBREAK, DT_SINGLELINE))
        Case "^" 'Centrado.
            lSuccess = DrawText(PFlex.hDC, ArrayCell(1), Len(ArrayCell(1)), MyRect, DT_VCENTER Or DT_CENTER Or IIf(LSeeAllText = True, DT_WORDBREAK, DT_SINGLELINE))
        End Select

'Dibujo el marco intermitente de seleccion de celda.
        If Fila = LRow And Columna = LCol And FocusRect <> None Then
            Select Case FocusRect
            Case Light
                PFlex.ForeColor = vbBlack
                MyRect.Left = X + 1
                MyRect.Right = X + AnchoColumna
                MyRect.Top = Y + 1
                MyRect.Bottom = Y + (VirtualHeight / 15)
                DrawFocusRect PFlex.hDC, MyRect
            Case Heavy
                PFlex.FillStyle = 1
                PFlex.DrawWidth = 2
                PFlex.ForeColor = LSelBackColor
                Rectangle PFlex.hDC, X + 2, Y + 2, X + AnchoColumna, Y + (VirtualHeight / 15)
                PFlex.DrawWidth = 1
                PFlex.FillStyle = 0
            End Select
        End If

        X = X + AnchoColumna

        UltimaColumna = Columna

        If X > PFlex.ScaleWidth Then Exit For
    Next Columna
    X = 0
    Y = Y + (VirtualHeight / 15)

    UltimaFila = Fila

    If Y > PFlex.ScaleHeight Then Exit For
Next Fila
PFlex.ScaleMode = vbTwips
'Marco general que envuelve a las celdas.
PFlex.FillStyle = 1
PFlex.Line (0, VirtualHeight)-(AnchoTotal, AltoTotal), vbBlack, B
PFlex.FillStyle = 0
Adding = False 'No estamos añadiendo filas.
PFlex.FontName = UserControl.FontName
PFlex.Font.Size = UserControl.FontSize
PFlex.Font.Bold = UserControl.FontBold
End Function


Private Sub ScrollValueChanged()
Select Case ValorScroll
Case Ninguna
    Shape.Visible = False
Case Horizontal
    HsHorizontal.Top = PFlex.ScaleHeight - 225
    HsHorizontal.Width = PFlex.ScaleWidth - 465
    HsHorizontal.Left = 465
    CTop.Top = PFlex.ScaleHeight - 225
    CTop.Left = 15
    CLeft.Top = PFlex.ScaleHeight - 225
    CLeft.Left = 240
    Shape.Visible = False
Case Vertical
    VsVertical.Left = PFlex.ScaleWidth - 225
    VsVertical.Height = PFlex.ScaleHeight
    Shape.Visible = False
Case Ambas
    HsHorizontal.Top = PFlex.ScaleHeight - 225
    HsHorizontal.Width = PFlex.ScaleWidth - 690
    HsHorizontal.Left = 465
    CTop.Top = PFlex.ScaleHeight - 225
    CTop.Left = 15
    CLeft.Top = PFlex.ScaleHeight - 225
    CLeft.Left = 240
    VsVertical.Left = PFlex.ScaleWidth - 225
    VsVertical.Height = PFlex.ScaleHeight - 225
    Shape.Top = PFlex.ScaleHeight - 225
    Shape.Left = PFlex.ScaleWidth - 225
    Shape.Visible = True

End Select
End Sub

Public Property Get Scrolls() As TipoScroll
Attribute Scrolls.VB_Description = "Devuelve o establece que  barras de desplazamiento se verán."
Attribute Scrolls.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
Scrolls = ValorScroll
End Property

Public Property Let Scrolls(ByVal vNewValue As TipoScroll)
Select Case vNewValue
Case Ninguna
    HsHorizontal.Visible = False
    CTop.Visible = False
    CLeft.Visible = False
    VsVertical.Visible = False
    ValorScroll = vNewValue

Case Horizontal
    HsHorizontal.Visible = True
    CTop.Visible = True
    CLeft.Visible = True
    VsVertical.Visible = False
    ValorScroll = vNewValue

Case Vertical
    HsHorizontal.Visible = False
    CTop.Visible = False
    CLeft.Visible = False
    VsVertical.Visible = True
    ValorScroll = vNewValue

Case Ambas
    HsHorizontal.Visible = True
    CTop.Visible = True
    CLeft.Visible = True
    VsVertical.Visible = True
    ValorScroll = vNewValue

End Select
Call ScrollValueChanged
PropertyChanged "Scrolls"
End Property

Private Sub CLeft_Click()
If HsHorizontal.Value <> HsHorizontal.Max Then HsHorizontal.Value = HsHorizontal.Max
End Sub

Private Sub CLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PFlex.SetFocus
End Sub

Private Sub CTop_Click()
If VsVertical.Value <> VsVertical.Max Then VsVertical.Value = VsVertical.Max
End Sub

Private Sub CTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PFlex.SetFocus
End Sub

Private Sub HsHorizontal_Change()
Call CalcularAnchoTotal
Call DibujarCeldas
PFlex.Refresh
End Sub

Private Sub HsHorizontal_GotFocus()
PFlex.SetFocus
End Sub

Private Sub HsHorizontal_Scroll()
If ScrollTrack = True Then
    Call CalcularAnchoTotal
    Call DibujarCeldas
    PFlex.Refresh
End If
End Sub

Private Sub PFlex_Click()
RaiseEvent Click
End Sub

Private Sub PFlex_DblClick()
RaiseEvent DblClick
End Sub

Private Sub PFlex_KeyDown(KeyCode As Integer, Shift As Integer)
Dim NuevoValor As Integer, i As Integer, Ancho As Integer
If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyLeft 'Tecla izquierda.
        Row = Row
        If LCol > 1 Then Col = Col - 1
    Case vbKeyRight 'Tecla derecha.
        Row = Row
        If LCol < LCols Then Col = Col + 1
    Case vbKeyUp 'Tecla arriba.
        Col = Col
        If LRow > 1 Then Row = Row - 1
    Case vbKeyDown 'Tecla abajo.
        Col = Col
        If LRow < LRows Then Row = Row + 1
    Case vbKeyPageUp
        Col = Col
        If LRow - 10 >= 1 Then Row = Row - 10 Else Row = 1
    Case vbKeyPageDown
        Col = Col
        If LRow + 10 <= LRows Then Row = Row + 10 Else Row = Rows
    End Select
End If

If KeyCode = vbKeyLeft And Shift = 1 Then
    If ColeccionColWidth.item(LCol) >= 120 Then
        NuevoValor = ColeccionColWidth.item(LCol)
        ColeccionColWidth.Remove (LCol)
        If LCol - 1 = ColeccionColWidth.Count Then
            ColeccionColWidth.Add NuevoValor - 30
        Else
            ColeccionColWidth.Add NuevoValor - 30, CStr(LCol), LCol
        End If
        Call CalcularAnchoTotal
        Call DibujarCeldas
        PFlex.Refresh
    End If
End If

If KeyCode = vbKeyRight And Shift = 1 Then
    NuevoValor = ColeccionColWidth.item(LCol)
    ColeccionColWidth.Remove (LCol)
    If LCol - 1 = ColeccionColWidth.Count Then
        ColeccionColWidth.Add NuevoValor + 30
    Else
        ColeccionColWidth.Add NuevoValor + 30, CStr(LCol), LCol
    End If
    Call CalcularAnchoTotal
    Call DibujarCeldas
    PFlex.Refresh
End If

If KeyCode = vbKeyUp And Shift = 1 Then
    If VirtualHeight >= 270 Then
        VirtualHeight = VirtualHeight - 30
        Call CalcularAltoTotal
        Call DibujarCeldas
        PFlex.Refresh
    End If
End If

If KeyCode = vbKeyDown And Shift = 1 Then
    If VirtualHeight <= 960 Then
        VirtualHeight = VirtualHeight + 30
        Call CalcularAltoTotal
        Call DibujarCeldas
        PFlex.Refresh
    End If
End If
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub PFlex_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub


Private Sub PFlex_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub PFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, Ancho As Integer, AnteriorRow As Integer, AnteriorCol As Integer
AnteriorRow = LRow
AnteriorCol = LCol
If Button = 1 Then
'Calculo mouserow y mousecol.
    Call CalcularMouseRowCol(X, Y)
    ClickRowCol = True

    If AnteriorRow <> MouseRow Or AnteriorCol <> MouseCol Then RaiseEvent LeaveCell
    Row = MouseRow
    Col = MouseCol
    If AnteriorRow <> MouseRow Or AnteriorCol <> MouseCol Then RaiseEvent EnterCell

    ClickRowCol = False
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub PFlex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CalcularMouseRowCol(X, Y)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PFlex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Column As Integer
Static Ascendente As Boolean
'Cambio de posicion de columna.
If Button = 1 And Y <= VirtualHeight And AllowUserChangeColPos = True Then
    Call ChangeColPosition(Col, MouseCol)
End If
'Ordenacion por x columna.
If Button = 2 And Y <= VirtualHeight And AllowUserSortCol = True Then
    If MouseCol = Column Then
        If Ascendente = False Then Ascendente = True Else Ascendente = False
        Call Sorted(1, Rows, Column, Ascendente)
    Else
        Column = MouseCol
        Call Sorted(1, Rows, Column, True)
        Ascendente = True
    End If
    Column = MouseCol
End If
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub



Private Sub UserControl_InitProperties()
Scrolls = 3
Cols = 2
Rows = 1

RowHeight = 240
Titulo = ""
Backcolor = &H808080
CellsBackColor = vbWhite
GridColor = &HC0C0C0
ForeColor = vbBlack
TituloForeColor = vbBlack
TituloBackColor = &H8000000F
SelBackColor = &H8000000D
SelForeColor = vbWhite
ShowGrid = True
FocusRect = Light
SelRow = True
AllowUserChangeColPos = True
AllowUserSortCol = True

HsHorizontal.Min = 1
HsHorizontal.Max = Cols
HsHorizontal.Value = 1

VsVertical.Min = 1
VsVertical.Max = Rows
VsVertical.Value = 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set PFlex.Font = PropBag.ReadProperty("Font", Ambient.Font)
Scrolls = PropBag.ReadProperty("Scrolls", 3)
RowHeight = PropBag.ReadProperty("RowHeight", 240)
Cols = PropBag.ReadProperty("Cols", 1)
Rows = PropBag.ReadProperty("Rows", 1)
Titulo = PropBag.ReadProperty("Titulo", "")
MouseRow = PropBag.ReadProperty("MouseRow", 0)
MouseCol = PropBag.ReadProperty("Mousecol", 0)
If LRows = 0 Then
    Row = PropBag.ReadProperty("Row", 0)
    Col = PropBag.ReadProperty("col", 0)
    AntCol = 0
    AntRow = 0
Else
    Row = PropBag.ReadProperty("Row", 1)
    Col = PropBag.ReadProperty("Col", 1)
    AntCol = 1
    AntRow = 1
End If
Backcolor = PropBag.ReadProperty("Backcolor", &H808080)
CellsBackColor = PropBag.ReadProperty("CellsBackColor", vbWhite)
GridColor = PropBag.ReadProperty("Gridcolor", &HC0C0C0)
ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
TituloForeColor = PropBag.ReadProperty("TituloForeColor", vbBlack)
TituloBackColor = PropBag.ReadProperty("TituloBackColor", &H8000000F)
SelBackColor = PropBag.ReadProperty("SelBackColor", &H8000000D)
SelForeColor = PropBag.ReadProperty("SelForeColor", vbWhite)
ShowGrid = PropBag.ReadProperty("ShowGrid", True)
SelRow = PropBag.ReadProperty("SelRow", True)
AutoScrolls = PropBag.ReadProperty("AutoScrolls", False)
FocusRect = PropBag.ReadProperty("FocusRect", Light)
ScrollTrack = PropBag.ReadProperty("ScrollTrack", False)
SeeAllText = PropBag.ReadProperty("SeeAllText", False)
AllowUserChangeColPos = PropBag.ReadProperty("AllowUserChangeColPos", True)
AllowUserSortCol = PropBag.ReadProperty("AllowUserSortCol", True)

HsHorizontal.Min = 1
If Cols = 0 Then HsHorizontal.Max = 1 Else HsHorizontal.Max = Cols
HsHorizontal.Value = 1

VsVertical.Min = 1
If Rows = 0 Then VsVertical.Max = 1 Else VsVertical.Max = Rows
VsVertical.Value = 1

'Call DibujarCeldas
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
PFlex.Height = UserControl.Height
PFlex.Width = UserControl.Width
Call CalcularAnchoTotal
Call CalcularAltoTotal
Call ScrollValueChanged

Call DibujarCeldas
End Sub

Private Sub UserControl_Terminate()
Set ColeccionTitulo = Nothing
Set ColeccionColWidth = Nothing
Set ColeccionAlineacion = Nothing
If LRows > 0 Then Set ArrayRows = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

PropBag.WriteProperty "Scrolls", Scrolls, 3
Call PropBag.WriteProperty("Font", PFlex.Font, Ambient.Font)
PropBag.WriteProperty "RowHeight", RowHeight, 240
PropBag.WriteProperty "Rows", Rows, 1
PropBag.WriteProperty "Cols", Cols, 1
PropBag.WriteProperty "Titulo", Titulo, ""
PropBag.WriteProperty "MouseRow", MouseRow, 0
PropBag.WriteProperty "MouseCol", MouseCol, 0
PropBag.WriteProperty "Row", Row, 1
PropBag.WriteProperty "Col", Col, 1
PropBag.WriteProperty "Backcolor", Backcolor, &H808080
PropBag.WriteProperty "CellsBackColor", CellsBackColor, vbWhite
PropBag.WriteProperty "GridColor", GridColor, &HC0C0C0
PropBag.WriteProperty "ForeColor", ForeColor, vbBlack
PropBag.WriteProperty "TituloForeColor", TituloForeColor, vbBlack
PropBag.WriteProperty "TituloBackColor", TituloBackColor, &H8000000F
PropBag.WriteProperty "SelBackColor", SelBackColor, &H8000000D
PropBag.WriteProperty "SelForeColor", SelForeColor, vbWhite
PropBag.WriteProperty "ShowGrid", ShowGrid, True
PropBag.WriteProperty "SelRow", SelRow, True
PropBag.WriteProperty "AutoScrolls", AutoScrolls, False
PropBag.WriteProperty "FocusRect", FocusRect, Light
PropBag.WriteProperty "ScrollTrack", ScrollTrack, False
PropBag.WriteProperty "SeeAllText", SeeAllText, False
PropBag.WriteProperty "AllowUserChangeColPos", AllowUserChangeColPos, True
PropBag.WriteProperty "AllowUserSortCol", AllowUserSortCol, True

End Sub

Public Property Get Rows() As Integer
Attribute Rows.VB_Description = "Devuelve o establece el número de filas."
Rows = LRows
End Property

Public Property Let Rows(ByVal vNewRow As Integer)
Dim i As Integer, Dato As String, e As Integer, Mayor As Long
Dim AntRows As Integer

AntRows = LRows
If vNewRow = 0 Then
    Row = 0
    Col = 0
    LRows = vNewRow
    If AntRows > 0 Then Set ArrayRows = Nothing
    RowsCompletas = False
End If

LRows = vNewRow

If vNewRow = 0 Then VsVertical.Max = 1 Else VsVertical.Max = vNewRow

'Cada vez que hay un cambio en las filas, recalculo el alto total para
'repintar el cuadro de fondo.
Call CalcularAltoTotal

If Adding = True Then
    If RowsCompletas = False Then
        Call DibujarCeldas
        PFlex.Refresh
    Else
        Adding = False
    End If
Else
'aqui meter codigo si hay cambio en las rows para añadir a arrayrows.
    If LRows <> 0 Then
        If AntRows = 0 Then
            ReDim ArrayRows(LRows)
            For i = 1 To LRows
                ReDim ArrayCols(LCols)
                For e = 1 To LCols
                    ReDim ArrayCell(10)
                    ArrayCols(e) = ArrayCell
                Next e
                ArrayRows(i) = ArrayCols
            Next i
        Else
            Mayor = UBound(ArrayRows)
            ReDim Preserve ArrayRows(LRows)
            If LRows > Mayor Then
                For i = Mayor + 1 To LRows
                    ReDim ArrayCols(LCols)
                    For e = 1 To LCols
                        ReDim ArrayCell(10)
                        ArrayCols(e) = ArrayCell
                    Next e
                    ArrayRows(i) = ArrayCols
                Next i
            Else
'Si la fila anterior es superior a la ultima nueva fila.
'cambia el valor de row al valor ultimo es decir rows.
                If Row > Rows Then Row = Rows

            End If
        End If
    End If
    Call DibujarCeldas
End If
PropertyChanged "Rows"
End Property

Public Property Get Cols() As Integer
Attribute Cols.VB_Description = "Devuelve o establece el número de columnas."
Cols = LCols
End Property

Public Property Let Cols(ByVal vNewCols As Integer)
Dim i As Integer
Select Case vNewCols
Case Is > LCols
    If ColeccionColWidth.Count >= 0 And ColeccionColWidth.Count <> vNewCols Then
        For i = 1 To (vNewCols - LCols)
            ColeccionColWidth.Add LColWidth
            ColeccionAlineacion.Add "<"
            ColeccionTitulo.Add " "
        Next i
    End If
'redimensiono los arraysrow al numero de columnas, si rows > 0.
    If LRows > 0 Then
        If LCols = 0 Then
            ReDim ArrayRows(LRows)
        End If

        For i = 1 To UBound(ArrayRows)
            ArrayCols = ArrayRows(i)

            If LCols = 0 Then
                ReDim ArrayCols(vNewCols)
                For e = 1 To vNewCols
                    ReDim ArrayCell(10)
                    ArrayCols(e) = ArrayCell
                Next e
                ArrayRows(i) = ArrayCols
            Else
                ReDim Preserve ArrayCols(vNewCols)
                For e = LCols + 1 To vNewCols
                    ReDim ArrayCell(10)
                    ArrayCols(e) = ArrayCell
                Next e
                ArrayRows(i) = ArrayCols
            End If
            ArrayRows(i) = ArrayCols
        Next i
    End If
Case Is < LCols And vNewCols <> 0
    For i = vNewCols + 1 To ColeccionTitulo.Count
        ColeccionTitulo.Remove (ColeccionTitulo.Count)
        ColeccionColWidth.Remove (ColeccionColWidth.Count)
        ColeccionAlineacion.Remove (ColeccionAlineacion.Count)
    Next i
'redimensiono los arraysrow al numero de columnas.
    For i = 1 To UBound(ArrayRows)
        ArrayCols = ArrayRows(i)
        ReDim Preserve ArrayCols(vNewCols)
        ArrayRows(i) = ArrayCols
    Next i
'Si la columna anterior es superior a la ultima nueva columna.
'cambia el valor de col al valor ultimo es decir cols.
    LCols = vNewCols
    If Col > vNewCols Then Col = vNewCols
Case Is = 0
    Rows = 0
    Set ColeccionTitulo = Nothing
    Set ColeccionColWidth = Nothing
    Set ColeccionAlineacion = Nothing
End Select

LCols = vNewCols

HsHorizontal.Max = vNewCols

Call CalcularAnchoTotal

Call DibujarCeldas

PropertyChanged "Cols"
End Property

Public Property Get RowHeight() As Integer
Attribute RowHeight.VB_Description = "Devuelve o establece el alto de las filas en Twips. (Múltiplos de 15). Limitado a mínimo 240 y máximo  990."
RowHeight = VirtualHeight
End Property

Public Property Let RowHeight(ByVal vNewHeight As Integer)
Dim THeight As Integer, TtHeight As Integer
If vNewHeight < 991 Then
    If vNewHeight < 240 Then
        vNewHeight = 240
    Else
        THeight = vNewHeight Mod 15
        TtHeight = ((vNewHeight) - THeight)
        vNewHeight = TtHeight
        VirtualHeight = vNewHeight
    End If
Else
    VirtualHeight = 990
End If
Call CalcularAltoTotal
Call DibujarCeldas
PropertyChanged "RowHeight"
End Property

Public Property Get Titulo() As String
Attribute Titulo.VB_Description = "Determina el título del grid. | =cambio de columna, < =alineación a la izquierda, > =alineación a la derecha y ^ =centrado."
Titulo = VirtualTitulo
End Property

Public Property Let Titulo(ByVal vNewTitulo As String)
VirtualTitulo = vNewTitulo
Call CargarTitulo
PropertyChanged "Titulo"
End Property

Private Function CargarTitulo()
Dim TotalColumnas As Integer, Mipos As Long, Inicio As Long
Dim LastCols As Integer, TrozoTitulo As String, Alineacion As String, Trozo1 As String
LastCols = LCols
Mipos = 1
Inicio = 1
TotalColumnas = 1
'Número de columnas debido a titulo.
If VirtualTitulo <> "" Then
    Set ColeccionTitulo = Nothing
    Set ColeccionColWidth = Nothing
    Set ColeccionAlineacion = Nothing
    Do
        Mipos = InStr(Inicio, VirtualTitulo, "|")
        If Mipos = 0 Then
            TotalColumnas = 1
            Select Case Mid(VirtualTitulo, 1, 1)
            Case "<", ">", "^"
                TrozoTitulo = Mid(VirtualTitulo, 1, Len(VirtualTitulo) - 1)
                Alineacion = Mid(VirtualTitulo, 1, 1)
            Case Else
                TrozoTitulo = VirtualTitulo
                Alineacion = "<"
            End Select
            ColeccionTitulo.Add Trim(TrozoTitulo)
            ColeccionAlineacion.Add Alineacion

'calculo el ancho de la columna.
'If TextWidth(TrozoTitulo) > LColWidth Then
            ColeccionColWidth.Add TextWidth(TrozoTitulo) + 75
'Else
'    ColeccionColWidth.Add LColWidth
'End If

            Exit Do
        Else
            Trozo1 = Mid(VirtualTitulo, Inicio, (Mipos - 1) - (Inicio - 1))
            Select Case Mid(Trozo1, 1, 1)
            Case "<", ">", "^"
                TrozoTitulo = Mid(Trozo1, 2, Len(Trozo1) - 1)
                Alineacion = Mid(Trozo1, 1, 1)
            Case Else
                TrozoTitulo = Trozo1
                Alineacion = "<"
            End Select
            ColeccionTitulo.Add Trim(TrozoTitulo)
            ColeccionAlineacion.Add Alineacion

'calculo el ancho de la columna.
'If TextWidth(TrozoTitulo) > LColWidth Then
            ColeccionColWidth.Add TextWidth(TrozoTitulo) + 75
'Else
'    ColeccionColWidth.Add LColWidth
'End If
            TotalColumnas = TotalColumnas + 1
            Inicio = Mipos + 1
            Mipos = InStr(Inicio, VirtualTitulo, "|")
            If Mipos = 0 Then
                Trozo1 = Mid(VirtualTitulo, Inicio, Len(VirtualTitulo) - (Inicio - 1))
                Select Case Mid(Trozo1, 1, 1)
                Case "<", ">", "^"
                    TrozoTitulo = Mid(Trozo1, 2, Len(Trozo1) - 1)
                    Alineacion = Mid(Trozo1, 1, 1)
                Case Else
                    TrozoTitulo = Trozo1
                    Alineacion = "<"
                End Select
                ColeccionTitulo.Add Trim(TrozoTitulo)
                ColeccionAlineacion.Add Alineacion

'calculo el ancho de la columna.
'If TextWidth(TrozoTitulo) > LColWidth Then
                ColeccionColWidth.Add TextWidth(TrozoTitulo) + 75
'Else
'    ColeccionColWidth.Add LColWidth
'End If
                Exit Do
            End If

        End If
    Loop
    Cols = TotalColumnas 'If LastCols > TotalColumnas Then Cols = LastCols Else Cols = TotalColumnas
Else
    Set ColeccionTitulo = Nothing
    Cols = LastCols
End If

End Function

Private Sub VsVertical_Change()
Call CalcularAltoTotal
Call DibujarCeldas
PFlex.Refresh
End Sub

Public Function Clear()
Dim i As Integer, e As Integer

Set ColeccionTitulo = Nothing
Set ColeccionColWidth = Nothing
Set ColeccionAlineacion = Nothing
If LRows > 0 Then
    Set ArrayRows = Nothing
'Redimensiono el array que contiene las filas.
    ReDim ArrayRows(LRows)
    For i = 1 To LRows
'Redimensiono el array de fila al numero de columnas.
        ReDim ArrayCols(LCols)
'Asigno un valor vacio a cada propiedad por celda.
        For e = 1 To LCols
'Redimensiono el array de las propiedades.
            ReDim ArrayCell(10)
'Asigno las propiedades a cada columna.
            ArrayCols(e) = ArrayCell
        Next e
'asigno el valor de la fila.
        ArrayRows(i) = ArrayCols
    Next i
End If

For e = 1 To LCols
    ColeccionColWidth.Add LColWidth
    ColeccionAlineacion.Add "<"
Next e

Call CalcularAnchoTotal
Call DibujarCeldas

End Function

Public Property Get MouseRow() As Integer
Attribute MouseRow.VB_MemberFlags = "400"
MouseRow = VirtualMouseRow
End Property

Public Property Let MouseRow(ByVal vNewMouseRow As Integer)
VirtualMouseRow = vNewMouseRow
PropertyChanged "MouseRow"
End Property

Public Property Get MouseCol() As Integer
Attribute MouseCol.VB_MemberFlags = "400"
MouseCol = VirtualMouseCol
End Property

Public Property Let MouseCol(ByVal vNewMouseCol As Integer)
VirtualMouseCol = vNewMouseCol
PropertyChanged "MouseCol"
End Property

Public Function TopRow(Optional ByVal Index As Integer)
If Index = 0 Then
    VsVertical.Value = VsVertical.Max
Else
    VsVertical.Value = Index
End If
End Function
Public Function LeftCol(Optional ByVal Index As Integer)
If Index = 0 Then
    HsHorizontal.Value = HsHorizontal.Max
Else
    HsHorizontal.Value = Index
End If
End Function

Private Function Datos(Row As Integer, Col As Integer) As Variant
If LRow <> 0 And LCol <> 0 Then
    ArrayCols = ArrayRows(Row)
    ArrayCell = ArrayCols(Col)
    Datos = ArrayCell(1)
    Erase ArrayCell
    Erase ArrayCols
End If
End Function

Public Property Get TextMatrix(ByVal Row As Integer, ByVal Col As Integer) As String
TextMatrix = Datos(Row, Col)
End Property

Public Property Let TextMatrix(ByVal Row As Integer, ByVal Col As Integer, ByVal vNewValue As String)
'Obtengo los datos de la fila.
ArrayCols = ArrayRows(Row)
'Obtengo los datos de la celda de la columna especificada.
ArrayCell = ArrayCols(Col)
'inserto en la posicion 1 el nuevo valor.
ArrayCell(1) = vNewValue
'lleno los datos con el nuevo valor.
ArrayCols(Col) = ArrayCell
'lleno los datos de la fila con el nuevo valor.
ArrayRows(Row) = ArrayCols
'elimino el array.
Erase ArrayCols
'Redibujo.
If Row >= VsVertical.Value And Row <= UltimaFila And Col >= HsHorizontal.Value And Col <= UltimaColumna Then Call DibujarCelda(Row, Col)
'Call DibujarCeldas

End Property

Private Function DibujarTitulo()
Dim Fila As Integer, Columna As Integer, X As Long, Y As Long, AnchoColumna As Integer
Dim ValorHorizontal As Integer, ValorVertical As Integer
Dim lSuccess As Long, MyRect As RECT
PFlex.FillStyle = 1
ValorHorizontal = IIf(HsHorizontal.Value = 0, 1, HsHorizontal.Value)
ValorVertical = IIf(VsVertical.Value = 0, 1, VsVertical.Value)
'Dibujo las celdas del titulo.
PFlex.Cls
PFlex.ForeColor = LTituloForeColor
For Columna = ValorHorizontal To LCols
    If ColeccionColWidth.Count > 0 Then
        If ColeccionColWidth.item(Columna) = "" Then
            AnchoColumna = LColWidth
        Else
            AnchoColumna = ColeccionColWidth.item(Columna)
        End If
    Else
        AnchoColumna = LColWidth
    End If
    PFlex.Line (X, Y)-(X + AnchoColumna, Y + VirtualHeight), LTituloBackColor, BF
    If LShowGrid = True Then
        PFlex.Line (X + 15, 15)-(X + AnchoColumna, Y + VirtualHeight), vbWhite, B
        PFlex.Line (X, Y)-(X + AnchoColumna, Y + VirtualHeight), vbBlack, B
    End If
'cargo el titulo
    If VirtualTitulo <> "" Then
        PFlex.ScaleMode = vbPixels
        MyRect.Left = (X + 45) / 15
        MyRect.Right = (X + AnchoColumna - 30) / 15
        MyRect.Top = (Y / 15) + 1
        MyRect.Bottom = (Y + VirtualHeight) / 15
        Select Case ColeccionAlineacion.item(Columna)
        Case "<"
            If ColeccionTitulo.Count >= LCols Then
                lSuccess = DrawText(PFlex.hDC, ColeccionTitulo.item(Columna), Len(ColeccionTitulo.item(Columna)), MyRect, DT_VCENTER Or DT_LEFT Or DT_SINGLELINE)
            End If
        Case ">"
            lSuccess = DrawText(PFlex.hDC, ColeccionTitulo.item(Columna), Len(ColeccionTitulo.item(Columna)), MyRect, DT_VCENTER Or DT_RIGHT Or DT_SINGLELINE)
        Case "^"
            lSuccess = DrawText(PFlex.hDC, ColeccionTitulo.item(Columna), Len(ColeccionTitulo.item(Columna)), MyRect, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE)
        End Select
        PFlex.ScaleMode = vbTwips
    End If
    X = X + AnchoColumna
    If X > PFlex.ScaleWidth Then Exit For

Next Columna
If LShowGrid = False Then
    PFlex.Line (0, 0)-(AnchoTotal, VirtualHeight), vbBlack, B
End If
PFlex.FillStyle = 0
End Function

Public Property Get Row() As Integer
Attribute Row.VB_Description = "Devuelve o establece la fila especificada."
Attribute Row.VB_MemberFlags = "400"
Row = LRow
End Property

Public Property Let Row(ByVal vNewValue As Integer)
Dim i As Integer, Alto As Integer
AntRow = LRow
'Solo entra en la propiedad si el nuevo valor de fila es diferente al anterior.
If vNewValue <> LRow Then
'Nueva fila es menor que antigua fila.
    If vNewValue < LRow Then
        If ClickRowCol = False Then RaiseEvent LeaveCell
        LRow = vNewValue
        If ClickRowCol = False Then RaiseEvent EnterCell
        Select Case LRow
'Dentro de la cuadricula o por debajo del valor vertical.
        Case Is >= VsVertical.Value
            If LRow < UltimaFila And AntRow <> 0 Then
                Alto = IIf(HsHorizontal.Visible = True, 255, 0)
                If ((LRow + 2 - VsVertical.Value) * VirtualHeight) + Alto > PFlex.ScaleHeight And VsVertical.Value <> VsVertical.Max Then
                    VsVertical.Value = VsVertical.Value + 1
                Else
'Dibujo la antigua celda si esta en la cuadricula.
                    If AntRow >= VsVertical.Value And AntRow <= UltimaFila And AntRow <> 0 And AntRow <= Rows Then
                        For i = HsHorizontal.Value To UltimaColumna
                            Call DibujarCelda(AntRow, i)
                        Next i
                    End If
                    For i = HsHorizontal.Value To UltimaColumna
'Dibujo la nueva celda que sí está en la cuadricula.
                        Call DibujarCelda(vNewValue, i)
                    Next i
                End If
            Else
'La celda nueva es superior a la cuadricula. _
 En este caso la fila marcada será la ultima que _
 podamos ver en la cuadricula.
                Alto = IIf(HsHorizontal.Visible = True, 255, 0)
                For i = 1 To LRows
                    If ((i + 1) * VirtualHeight) + Alto > PFlex.ScaleHeight Then
                        VsVertical.Value = LRow - i + 2
                        Exit For
                    End If
                Next i
            End If
'Por encima del valor vertical.
        Case Is < VsVertical.Value
            If LRow <> 0 Then VsVertical.Value = LRow
        End Select
        PFlex.Refresh
    End If
'Nueva fila es mayor que antigua fila.
    If vNewValue > LRow Then
        If ClickRowCol = False Then RaiseEvent LeaveCell
        LRow = vNewValue
        If ClickRowCol = False Then RaiseEvent EnterCell
        Select Case LRow
'La celda nueva es mayor= que el valor de vsvertical.
        Case Is >= VsVertical.Value
'La celda nueva esta dentro de la cuadricula.
            If LRow <= UltimaFila Then
                Alto = IIf(HsHorizontal.Visible = True, 255, 0)
                If ((LRow + 2 - VsVertical.Value) * VirtualHeight) + Alto > PFlex.ScaleHeight And VsVertical.Value <> VsVertical.Max Then
''''''''''''
                    For i = 1 To LRows
                        If ((i + 1) * VirtualHeight) + Alto > PFlex.ScaleHeight Then
                            VsVertical.Value = LRow - i + 2
                            Exit For
                        End If
                    Next i
'''''''''''''
                Else
'Dibujo la antigua celda si esta en la cuadricula.
                    If AntRow >= VsVertical.Value And AntRow <= UltimaFila And AntRow <> 0 Then
                        For i = HsHorizontal.Value To UltimaColumna
                            Call DibujarCelda(AntRow, i)
                        Next i
                    End If
                    For i = HsHorizontal.Value To UltimaColumna
'Dibujo la nueva celda que sí está en la cuadricula.
                        Call DibujarCelda(vNewValue, i)
                    Next i
                End If
'La celda nueva es superior a la cuadricula. _
 En este caso la fila marcada será la ultima que _
 podamos ver en la cuadricula.
            Else
                Alto = IIf(HsHorizontal.Visible = True, 255, 0)
                For i = 1 To LRows
                    If ((i + 1) * VirtualHeight) + Alto > PFlex.ScaleHeight Then
                        VsVertical.Value = LRow - i + 2
                        Exit For
                    End If
                Next i

            End If
'La celda nueva es menor que el valor de vsvertical.
        Case Is < VsVertical.Value
            VsVertical.Value = LRow
        End Select
        PFlex.Refresh
    End If
Else 'la fila es la misma.
    Select Case LRow
    Case Is < VsVertical.Value
        If LRow > 0 Then VsVertical.Value = LRow
    Case Else
        Alto = IIf(HsHorizontal.Visible = True, 255, 0)
        Dim a, c
        a = (LRow + 2 - VsVertical.Value)
        c = (a * VirtualHeight) + Alto
        If c > PFlex.ScaleHeight And VsVertical.Value <> VsVertical.Max Then
            For i = 1 To LRows
                If ((i + 1) * VirtualHeight) + Alto > PFlex.ScaleHeight Then
                    VsVertical.Value = LRow - i + 2
                    Exit For
                End If
            Next i
        End If
    End Select
End If
AntRow = LRow
PropertyChanged "Row"
End Property

Public Property Get Col() As Integer
Attribute Col.VB_Description = "Devuelve o establece la columna especificada."
Attribute Col.VB_MemberFlags = "400"
Col = LCol
End Property

Public Property Let Col(ByVal vNewValue As Integer)
Dim i As Integer, ValorHorizontal As Integer, e As Integer, Ancho As Long, Inicio As Integer
AntCol = LCol
'Solo entra en la propiedad si el nuevo valor de columna es diferente al anterior.
If vNewValue <> LCol Then
'Nueva columna es menor que antigua Columna.
    If vNewValue < LCol Then
        If ClickRowCol = False Then RaiseEvent LeaveCell
        LCol = vNewValue
        If ClickRowCol = False Then RaiseEvent EnterCell
        Select Case LCol
'Esta a la izquierda de la cuadricula.
        Case Is < HsHorizontal.Value
            If LCol <> 0 Then HsHorizontal.Value = LCol
'Está dentro de la cuadricula o la derecha de la cuadricula.
        Case Is >= HsHorizontal.Value
'Está dentro de la cuadricula.
            If LCol <= UltimaColumna Then
                Ancho = 0
                For i = HsHorizontal.Value To LCol
                    Ancho = IIf(ColeccionColWidth.Count > 0, Ancho + ColeccionColWidth.item(i), Ancho + LColWidth)
                Next i
                Ancho = IIf(VsVertical.Visible = True, Ancho + 255, Ancho)
                If Ancho >= PFlex.ScaleWidth And Ancho <> 0 Then
                    Inicio = HsHorizontal.Value
                    Do
                        Ancho = 0
                        For i = Inicio + ValorHorizontal To LCol
                            Ancho = IIf(ColeccionColWidth.Count > 0, Ancho + ColeccionColWidth.item(i), Ancho + LColWidth)
                        Next i
                        Ancho = IIf(VsVertical.Visible = True, Ancho + 255, Ancho)
                        If Ancho <= PFlex.ScaleWidth And Ancho <> 0 Then Exit Do
                        ValorHorizontal = ValorHorizontal + 1
                    Loop
                    If HsHorizontal.Value < HsHorizontal.Max Then HsHorizontal.Value = HsHorizontal.Value + ValorHorizontal
                Else
'Dibulo las celdas especificas.
                    If AntCol <= UltimaColumna And AntCol <> 0 And AntRow >= VsVertical.Value And AntCol <= LCols Then Call DibujarCelda(Row, AntCol)
                    If LRow >= VsVertical.Value Then Call DibujarCelda(Row, LCol)
                End If

'Está a la derecha de la cuadricula.
            Else
'Busco el ancho que entra dentro del pflex.
                Inicio = HsHorizontal.Value
                Do
                    Ancho = 0
                    For i = Inicio + ValorHorizontal To LCol
                        Ancho = IIf(ColeccionColWidth.Count > 0, Ancho + ColeccionColWidth.item(i), Ancho + LColWidth)
                    Next i
                    Ancho = IIf(VsVertical.Visible = True, Ancho + 255, Ancho)
                    If Ancho <= PFlex.ScaleWidth And Ancho <> 0 Then Exit Do
                    ValorHorizontal = ValorHorizontal + 1
                Loop
                If HsHorizontal.Value < HsHorizontal.Max Then HsHorizontal.Value = HsHorizontal.Value + ValorHorizontal
            End If
        End Select
        PFlex.Refresh
    End If
'Nueva Columna es mayor que antigua Columna.
    If vNewValue > LCol Then
        If ClickRowCol = False Then RaiseEvent LeaveCell
        LCol = vNewValue
        If ClickRowCol = False Then RaiseEvent EnterCell
        Select Case LCol
'Esta a la izquierda o es la primera de la cuadricula.
        Case Is <= HsHorizontal.Value
            If LCol < HsHorizontal.Value Then
                HsHorizontal.Value = LCol
            Else
                Call DibujarCelda(Row, LCol)
            End If
'Está dentro de la cuadricula o la derecha de la cuadricula.
        Case Is > HsHorizontal.Value
'Está dentro de la cuadricula.
            If LCol <= UltimaColumna Then
                Ancho = 0
                For i = HsHorizontal.Value To LCol
                    Ancho = IIf(ColeccionColWidth.Count > 0, Ancho + ColeccionColWidth.item(i), Ancho + LColWidth)
                Next i
                Ancho = IIf(VsVertical.Visible = True, Ancho + 255, Ancho)
                If Ancho >= PFlex.ScaleWidth And Ancho <> 0 Then
                    Inicio = HsHorizontal.Value
                    Do
                        Ancho = 0
                        For i = Inicio + ValorHorizontal To LCol
                            Ancho = IIf(ColeccionColWidth.Count > 0, Ancho + ColeccionColWidth.item(i), Ancho + LColWidth)
                        Next i
                        Ancho = IIf(VsVertical.Visible = True, Ancho + 255, Ancho)
                        If Ancho <= PFlex.ScaleWidth And Ancho <> 0 Then Exit Do
                        ValorHorizontal = ValorHorizontal + 1
                    Loop
                    If HsHorizontal.Value < HsHorizontal.Max Then
                        If HsHorizontal.Value + ValorHorizontal <= HsHorizontal.Max Then
                            HsHorizontal.Value = HsHorizontal.Value + ValorHorizontal
                        Else
                            HsHorizontal.Value = HsHorizontal.Value + 1
                        End If
                    End If

                Else
'Dibulo las celdas especificas.
                    If AntCol <= UltimaColumna And AntCol <> 0 And AntRow >= VsVertical.Value Then Call DibujarCelda(Row, AntCol)
                    If LRow >= VsVertical.Value Then Call DibujarCelda(Row, LCol)
                End If

'Está a la derecha de la cuadricula.
            Else
'Busco el ancho que entra dentro del pflex.
                Inicio = HsHorizontal.Value
                Do
                    Ancho = 0
                    For i = Inicio + ValorHorizontal To LCol
                        Ancho = IIf(ColeccionColWidth.Count > 0, Ancho + ColeccionColWidth.item(i), Ancho + LColWidth)
                    Next i
                    Ancho = IIf(VsVertical.Visible = True, Ancho + 255, Ancho)
                    If Ancho <= PFlex.ScaleWidth And Ancho <> 0 Then Exit Do
                    ValorHorizontal = ValorHorizontal + 1
                Loop
                If HsHorizontal.Value < HsHorizontal.Max Then HsHorizontal.Value = HsHorizontal.Value + ValorHorizontal
            End If
        End Select
        PFlex.Refresh
    End If
Else 'la columna es la misma
    Select Case LCol 'esta encima de la cuadricula.
    Case Is < HsHorizontal.Value
        If LCol > 0 Then HsHorizontal.Value = LCol
    Case Else 'esta en la cuadricula o por debajo.
        For i = HsHorizontal.Value To LCol
            If i <> 0 Then Ancho = IIf(ColeccionColWidth.Count > 0, Ancho + ColeccionColWidth.item(i), Ancho + LColWidth)
        Next i
        Ancho = IIf(VsVertical.Visible = True, Ancho + 255, Ancho)
        If Ancho >= PFlex.ScaleWidth And Ancho <> 0 Then
            Inicio = HsHorizontal.Value
            Do
                Ancho = 0
                For i = Inicio + ValorHorizontal To LCol
                    Ancho = IIf(ColeccionColWidth.Count > 0, Ancho + ColeccionColWidth.item(i), Ancho + LColWidth)
                Next i
                Ancho = IIf(VsVertical.Visible = True, Ancho + 255, Ancho)
                If Ancho <= PFlex.ScaleWidth And Ancho <> 0 Then Exit Do
                ValorHorizontal = ValorHorizontal + 1
            Loop
            If HsHorizontal.Value < HsHorizontal.Max Then HsHorizontal.Value = HsHorizontal.Value + ValorHorizontal
        End If
    End Select
End If
AntCol = LCol
PropertyChanged "Col"
End Property

Private Function CalcularAltoTotal()
Dim AltoParcial As Integer, Fila As Integer
AltoTotal = 0
For Fila = VsVertical.Value To LRows
    AltoParcial = AltoParcial + VirtualHeight
    If AltoParcial > PFlex.ScaleHeight Then
        AltoTotal = PFlex.ScaleHeight
'Compruebo autoscrolls.
        If LAutoScrolls = True Then
            If Scrolls <> Ambas And Scrolls <> Vertical Then
                If Scrolls = Horizontal Then
                    Scrolls = Ambas
                Else
                    Scrolls = Vertical
                End If
            End If
        End If
'fin comprobacion.
        Exit For
    Else
        AltoTotal = AltoParcial + VirtualHeight
    End If
Next Fila
End Function

Private Sub VsVertical_GotFocus()
PFlex.SetFocus
End Sub


Public Property Get Backcolor() As OLE_COLOR
Attribute Backcolor.VB_Description = "Devuelve o establece el color de fondo."
Attribute Backcolor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Backcolor = LBackcolor
End Property

Public Property Let Backcolor(ByVal vNewBackcolor As OLE_COLOR)
LBackcolor = vNewBackcolor
PFlex.Backcolor = vNewBackcolor
Call DibujarCeldas
PropertyChanged "Backcolor"
End Property

Public Property Get CellsBackColor() As OLE_COLOR
Attribute CellsBackColor.VB_Description = "Devuelve o establece el color de fondo de las celdas."
Attribute CellsBackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
CellsBackColor = LCellsBackColor
End Property

Public Property Let CellsBackColor(ByVal vNewCellscolor As OLE_COLOR)
LCellsBackColor = vNewCellscolor
Call DibujarCeldas
PropertyChanged "CellsBackColor"
End Property

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Devuelve o establece el color del marco de las celdas."
Attribute GridColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
GridColor = LGridColor
End Property

Public Property Let GridColor(ByVal vNewGridcolor As OLE_COLOR)
LGridColor = vNewGridcolor
Call DibujarCeldas
PropertyChanged "GridColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color del texto de las celdas."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
ForeColor = LForeColor
End Property

Public Property Let ForeColor(ByVal vNewForeColor As OLE_COLOR)
LForeColor = vNewForeColor
PFlex.ForeColor = vNewForeColor
Call DibujarCeldas
PropertyChanged "ForeColor"
End Property

Public Property Get TituloForeColor() As OLE_COLOR
Attribute TituloForeColor.VB_Description = "Devuelve o establece el color del texto del título."
Attribute TituloForeColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
TituloForeColor = LTituloForeColor
End Property

Public Property Let TituloForeColor(ByVal vNewTituloForeColor As OLE_COLOR)
LTituloForeColor = vNewTituloForeColor
Call DibujarCeldas
PropertyChanged "TituloForeColor"
End Property

Public Property Get TituloBackColor() As OLE_COLOR
Attribute TituloBackColor.VB_Description = "Devuelve o establece el color del fondo del título."
Attribute TituloBackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
TituloBackColor = LTituloBackColor
End Property

Public Property Let TituloBackColor(ByVal vNewTituloBackColor As OLE_COLOR)
LTituloBackColor = vNewTituloBackColor
Call DibujarCeldas
PropertyChanged "TituloBackColor"
End Property



Public Property Get SelBackColor() As OLE_COLOR
Attribute SelBackColor.VB_Description = "Devuelve o establece el color de fondo de la fila seleccionada."
Attribute SelBackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
SelBackColor = LSelBackColor
End Property

Public Property Let SelBackColor(ByVal vNewSelBackColor As OLE_COLOR)
LSelBackColor = vNewSelBackColor
Call DibujarCeldas
PropertyChanged "SelBackColor"
End Property

Public Property Get SelForeColor() As OLE_COLOR
Attribute SelForeColor.VB_Description = "Devuelve o establece el color del texto de la fila seleccionada."
Attribute SelForeColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
SelForeColor = LSelForeColor
End Property


Public Property Let SelForeColor(ByVal vNewSelForeColor As OLE_COLOR)
LSelForeColor = vNewSelForeColor
Call DibujarCeldas
PropertyChanged "SelForeColor"
End Property

Public Property Get ShowGrid() As Boolean
Attribute ShowGrid.VB_Description = "Devuelve o establece si se verá el marco de la celda o no."
ShowGrid = LShowGrid
End Property

Public Property Let ShowGrid(ByVal vNewShowGrid As Boolean)
LShowGrid = vNewShowGrid
Call DibujarCeldas
PropertyChanged "ShowGrid"
End Property

Public Function ChangeColPosition(ByVal Col As Integer, ByVal Position As Integer)
Dim MiArray As Variant, Miarray1 As Variant, i As Integer
Dim ArrayCol As Variant, ArrayPosicion As Variant

'la columna debe ser una columna válida, y la posición una posición válida _
 y la columna debe ser diferente de la posición y las columnas deben ser mayores de 1.
If Col <= LCols And Position <= LCols And Col <> Position And LCols > 1 Then
    Screen.MousePointer = 11
''''''''''''''''''''''''''''''''''''''''
'traspaso los datos de alineacion.
    ReDim MiArray(LCols)
    If ColeccionAlineacion.Count < LCols Then
        For i = ColeccionAlineacion.Count + 1 To LCols
            ColeccionAlineacion.Add "<"
        Next i
    End If
    MiArray(Col) = ColeccionAlineacion.item(Position)
    MiArray(Position) = ColeccionAlineacion.item(Col)
    For i = 1 To LCols
        If i <> Col And i <> Position Then
            MiArray(i) = ColeccionAlineacion.item(i)
        End If
    Next i
    Set ColeccionAlineacion = Nothing
    For i = 1 To LCols
        ColeccionAlineacion.Add MiArray(i)
    Next i
''''''''''''''''''''''''''''''''''''''''
'traspaso los datos de LColWidth.
    ReDim MiArray(LCols)
    If ColeccionColWidth.Count < LCols Then
        For i = ColeccionColWidth.Count + 1 To LCols
            ColeccionColWidth.Add LColWidth
        Next i
    End If
    MiArray(Col) = ColeccionColWidth.item(Position)
    MiArray(Position) = ColeccionColWidth.item(Col)
    For i = 1 To LCols
        If i <> Col And i <> Position Then
            MiArray(i) = ColeccionColWidth.item(i)
        End If
    Next i
    Set ColeccionColWidth = Nothing
    For i = 1 To LCols
        ColeccionColWidth.Add MiArray(i)
    Next i
''''''''''''''''''''''''''''''''''''''''
'traspaso los datos de texto.
    ReDim MiArray(LCols)
    If ColeccionTitulo.Count < LCols Then
        For i = ColeccionTitulo.Count + 1 To LCols
            ColeccionTitulo.Add " "
        Next i
    End If
    MiArray(Col) = ColeccionTitulo.item(Position)
    MiArray(Position) = ColeccionTitulo.item(Col)
    For i = 1 To LCols
        If i <> Col And i <> Position Then
            MiArray(i) = ColeccionTitulo.item(i)
        End If
    Next i
    Set ColeccionTitulo = Nothing
    For i = 1 To LCols
        ColeccionTitulo.Add MiArray(i)
    Next i
'ordeno las columna segun col y position.
    For i = 1 To UBound(ArrayRows)
        ArrayCols = ArrayRows(i)

        ArrayCol = ArrayCols(Col)
        ArrayPosicion = ArrayCols(Position)

        ArrayCols(Col) = ArrayPosicion
        ArrayCols(Position) = ArrayCol

        ArrayRows(i) = ArrayCols
    Next i
    Call DibujarCeldas
    Screen.MousePointer = 0
End If
End Function

Public Function Sorted(ByVal FirstRow As Integer, ByVal LastRow As Integer, ByVal Col As Integer, ByVal Ascending As Boolean)
Dim i As Integer, e As Integer, Colocado As Boolean
Dim Primero As String, Ultimo As String, Posicion As Integer, UltString As String
Dim CollLinea As New Collection
Dim ArrayColsComp As Variant, ArrayCellsComp As Variant
Dim ArrayColsComp1 As Variant, ArrayCellsComp1 As Variant
Dim MiComp As Integer, Micomp1 As Integer
'Deben heber minimo 2 filas
If LRows > 1 Then
    UserControl.MousePointer = 11
    For i = FirstRow To LastRow
        ArrayCols = ArrayRows(i)
        ArrayCell = ArrayCols(Col)
        If i = FirstRow Then
            CollLinea.Add ArrayRows(i)
            Primero = ArrayCell(1)
            Ultimo = ArrayCell(1)
            Posicion = 1
            UltString = ArrayCell(1)
        Else
            Select Case Ascending
            Case True
'Comparo con el último.
                MiComp = StrComp(ArrayCell(1), Ultimo, vbTextCompare)
                If MiComp = 1 Or MiComp = 0 Then
                    CollLinea.Add ArrayRows(i)
                    Ultimo = ArrayCell(1)
                    Posicion = CollLinea.Count
                    UltString = ArrayCell(1)
                    Colocado = True
                End If
'Comparo con el primero.
                If Colocado = False Then
                    MiComp = StrComp(ArrayCell(1), Primero, vbTextCompare)
                    If MiComp = -1 Or MiComp = 0 Then
                        CollLinea.Add ArrayRows(i), , 1
                        Primero = ArrayCell(1)
                        Posicion = 1
                        UltString = ArrayCell(1)
                        Colocado = True
                    End If
                End If
'Comparo con el ultimo colocado.
                If Colocado = False Then
                    MiComp = StrComp(ArrayCell(1), UltString, vbTextCompare)
                    Select Case MiComp
                    Case Is = 0
                        CollLinea.Add ArrayRows(i), , Posicion
                    Case Is = 1
                        For e = Posicion To CollLinea.Count   '-1 porque el ultimo ya se ha comprobado.
                            ArrayColsComp = CollLinea.item(e)
                            ArrayCellsComp = ArrayColsComp(Col)
                            MiComp = StrComp(ArrayCell(1), ArrayCellsComp(1), vbTextCompare)

                            ArrayColsComp1 = CollLinea.item(e + 1)
                            ArrayCellsComp1 = ArrayColsComp1(Col)
                            Micomp1 = StrComp(ArrayCell(1), ArrayCellsComp1(1), vbTextCompare)
                            If MiComp = 1 Or MiComp = 0 Then
                                If Micomp1 = -1 Or Micomp1 = 0 Then
                                    CollLinea.Add ArrayRows(i), , e + 1
                                    Posicion = e + 1
                                    UltString = ArrayCell(1)
                                    Exit For
                                End If
                            End If
                        Next e

                    Case Is = -1
                        For e = 1 To Posicion - 1 '-1 porque el ultimo ya se ha comprobado.
                            ArrayColsComp = CollLinea.item(e)
                            ArrayCellsComp = ArrayColsComp(Col)
                            MiComp = StrComp(ArrayCell(1), ArrayCellsComp(1), vbTextCompare)

                            ArrayColsComp1 = CollLinea.item(e + 1)
                            ArrayCellsComp1 = ArrayColsComp1(Col)
                            Micomp1 = StrComp(ArrayCell(1), ArrayCellsComp1(1), vbTextCompare)
                            If MiComp = 1 Or MiComp = 0 Then
                                If Micomp1 = -1 Or Micomp1 = 0 Then
                                    CollLinea.Add ArrayRows(i), , e + 1
                                    Posicion = e + 1
                                    UltString = ArrayCell(1)
                                    Exit For
                                End If
                            End If
                        Next e
                    End Select
                    Colocado = True
                End If
            Case False 'Descendente
'Comparo con el último.
                MiComp = StrComp(ArrayCell(1), Ultimo, vbTextCompare)
                If MiComp = -1 Or MiComp = 0 Then
                    CollLinea.Add ArrayRows(i)
                    Ultimo = ArrayCell(1)
                    Posicion = CollLinea.Count
                    UltString = ArrayCell(1)
                    Colocado = True
                End If
'Comparo con el primero.
                If Colocado = False Then
                    MiComp = StrComp(ArrayCell(1), Primero, vbTextCompare)
                    If MiComp = 1 Or MiComp = 0 Then
                        CollLinea.Add ArrayRows(i), , 1
                        Primero = ArrayCell(1)
                        Posicion = 1
                        UltString = ArrayCell(1)
                        Colocado = True
                    End If
                End If
                If Colocado = False Then
                    MiComp = StrComp(ArrayCell(1), UltString, vbTextCompare)
                    Select Case MiComp
                    Case Is = 0
                        CollLinea.Add ArrayRows(i), , Posicion
                    Case Is = 1
                        For e = 1 To Posicion - 1
                            ArrayColsComp = CollLinea.item(e)
                            ArrayCellsComp = ArrayColsComp(Col)
                            MiComp = StrComp(ArrayCell(1), ArrayCellsComp(1), vbTextCompare)

                            ArrayColsComp1 = CollLinea.item(e + 1)
                            ArrayCellsComp1 = ArrayColsComp1(Col)
                            Micomp1 = StrComp(ArrayCell(1), ArrayCellsComp1(1), vbTextCompare)
                            If MiComp = -1 Or MiComp = 0 And Micomp1 = 1 Or Micomp1 = 0 Then
                                CollLinea.Add ArrayRows(i), , e + 1
                                Posicion = e + 1
                                UltString = ArrayCell(1)
                                Exit For
                            End If
                        Next e

                    Case Is = -1
                        For e = Posicion To CollLinea.Count
                            ArrayColsComp = CollLinea.item(e)
                            ArrayCellsComp = ArrayColsComp(Col)
                            MiComp = StrComp(ArrayCell(1), ArrayCellsComp(1), vbTextCompare)

                            ArrayColsComp1 = CollLinea.item(e + 1)
                            ArrayCellsComp1 = ArrayColsComp1(Col)
                            Micomp1 = StrComp(ArrayCell(1), ArrayCellsComp1(1), vbTextCompare)
                            If MiComp = -1 Or MiComp = 0 And Micomp1 = 1 Or Micomp1 = 0 Then
                                CollLinea.Add ArrayRows(i), , e + 1
                                Posicion = e + 1
                                UltString = ArrayCell(1)
                                Exit For
                            End If
                        Next e
                    End Select
                    Colocado = True
                End If
            End Select
        End If
        Colocado = False
    Next i
    Dim FItem As Integer
    FItem = 1
    For i = FirstRow To LastRow
        ArrayRows(i) = CollLinea.item(FItem)
        FItem = FItem + 1
    Next i
    Call DibujarCeldas
    Set CollLinea = Nothing
    UserControl.MousePointer = 0
End If
End Function


Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "400"
Text = Datos(LRow, LCol)
End Property

Public Property Let Text(ByVal vNewText As String)
'Obtengo los datos de la fila.
ArrayCols = ArrayRows(LRow)
'Obtengo los datos de la celda de la columna especificada.
ArrayCell = ArrayCols(LCol)
'inserto en la posicion 1 el nuevo valor.
ArrayCell(1) = vNewText
'lleno los datos con el nuevo valor de columna.
ArrayCols(LCol) = ArrayCell
'lleno los datos de la fila con el nuevo valor de fila.
ArrayRows(LRow) = ArrayCols
'elimino el array.
Erase ArrayCols
'Redibujo.
If Row >= VsVertical.Value And Row <= UltimaFila And Col >= HsHorizontal.Value And Col <= UltimaColumna Then Call DibujarCelda(Row, Col)
'Call DibujarCeldas
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Fuente"
Attribute Font.VB_UserMemId = -512
Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal vnewfont As StdFont)

Set PFlex.Font = vnewfont
Set UserControl.Font = vnewfont

Call DibujarCeldas

PropertyChanged "Font"


End Property

Public Property Get CellForeColor(ByVal Row As Integer, ByVal Col As Integer) As OLE_COLOR
Attribute CellForeColor.VB_Description = "Devuelve o establece el color de fondo de una celda determinada."
Attribute CellForeColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute CellForeColor.VB_MemberFlags = "400"
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
CellForeColor = ArrayCell(3)
End Property

Public Property Get CellBackColor(ByVal Row As Integer, ByVal Col As Integer) As OLE_COLOR
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
CellBackColor = ArrayCell(2)
End Property
Public Property Get CellFontBold(ByVal Row As Integer, ByVal Col As Integer) As Boolean
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
CellFontBold = ArrayCell(6)
End Property

Public Property Get CellFontSize(ByVal Row As Integer, ByVal Col As Integer) As Single
Attribute CellFontSize.VB_Description = "Devuelve o establece el tamaño de la fuente de la celda especificada."
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
CellFontSize = ArrayCell(5)
End Property

Public Property Get CellFontName(ByVal Row As Integer, ByVal Col As Integer) As String
Attribute CellFontName.VB_Description = "Devuelve o establece el nombre de la fuente de la celda especificada."
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
CellFontName = ArrayCell(4)
End Property

Public Property Let CellBackColor(ByVal Row As Integer, ByVal Col As Integer, ByVal vNewValue As OLE_COLOR)
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
ArrayCell(2) = vNewValue

ArrayCols(Col) = ArrayCell
ArrayRows(Row) = ArrayCols
If Row >= VsVertical.Value And Row <= UltimaFila And Col >= HsHorizontal.Value And Col <= UltimaColumna Then
    Call DibujarCelda(Row, Col)
    PFlex.Refresh
End If
PropertyChanged "CellBackColor"

End Property
Public Property Let CellFontBold(ByVal Row As Integer, ByVal Col As Integer, ByVal vNewValue As Boolean)
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
ArrayCell(6) = vNewValue

ArrayCols(Col) = ArrayCell
ArrayRows(Row) = ArrayCols
If Row >= VsVertical.Value And Row <= UltimaFila And Col >= HsHorizontal.Value And Col <= UltimaColumna Then
    Call DibujarCelda(Row, Col)
    PFlex.Refresh
End If
PropertyChanged "CellFontBold"

End Property

Public Property Let CellFontSize(ByVal Row As Integer, ByVal Col As Integer, ByVal vNewValue As Single)
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
ArrayCell(5) = vNewValue

ArrayCols(Col) = ArrayCell
ArrayRows(Row) = ArrayCols
If Row >= VsVertical.Value And Row <= UltimaFila And Col >= HsHorizontal.Value And Col <= UltimaColumna Then
    Call DibujarCelda(Row, Col)
    PFlex.Refresh
End If
PropertyChanged "CellFontSize"

End Property

Public Property Let CellFontName(ByVal Row As Integer, ByVal Col As Integer, ByVal vNewValue As String)
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
ArrayCell(4) = vNewValue

ArrayCols(Col) = ArrayCell
ArrayRows(Row) = ArrayCols
If Row >= VsVertical.Value And Row <= UltimaFila And Col >= HsHorizontal.Value And Col <= UltimaColumna Then
    Call DibujarCelda(Row, Col)
    PFlex.Refresh
End If
PropertyChanged "CellFontName"

End Property

Public Property Let CellForeColor(ByVal Row As Integer, ByVal Col As Integer, ByVal vNewValue As OLE_COLOR)
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)
ArrayCell(3) = vNewValue

ArrayCols(Col) = ArrayCell
ArrayRows(Row) = ArrayCols
If Row >= VsVertical.Value And Row <= UltimaFila And Col >= HsHorizontal.Value And Col <= UltimaColumna Then
    Call DibujarCelda(Row, Col)
    PFlex.Refresh
End If
PropertyChanged "CellForeColor"

End Property

Private Function DibujarCelda(ByVal Row As Integer, ByVal Col As Integer)
Dim X As Long, Y As Long, AnchoColumna As Integer
Dim i As Integer, Negrita As Boolean
Dim lSuccess As Long, MyRect As RECT

PFlex.ScaleMode = vbPixels
ArrayCols = ArrayRows(Row)
ArrayCell = ArrayCols(Col)

X = 0

For i = HsHorizontal.Value To Col - 1
    X = X + IIf(ColeccionColWidth.Count > 0, ColeccionColWidth.item(i) / 15, LColWidth / 15)
Next i
'Ancho de la columna.
AnchoColumna = IIf(ColeccionColWidth.Count > 0, ColeccionColWidth.item(Col) / 15, LColWidth / 15)

Y = ((Row - VsVertical.Value) + 1) * (VirtualHeight / 15)

'Fuente de la celda.
If ArrayCell(4) <> "" Then PFlex.Font = ArrayCell(4)
'Tamaño de la fuente de la celda.
If ArrayCell(5) <> "" Then PFlex.Font.Size = ArrayCell(5)
'Bold de la fuente
Negrita = ArrayCell(6)
PFlex.Font.Bold = Negrita
'Color del marco de la celda.
PFlex.ForeColor = IIf(ShowGrid = True, LGridColor, LCellsBackColor)
'Color de la celda.
PFlex.FillColor = IIf(Row = LRow And Col <> LCol And LSelRow = True, SelBackColor, IIf(ArrayCell(2) <> "", ArrayCell(2), LCellsBackColor))
'Dibujo la celda.
Rectangle PFlex.hDC, X, Y, X + 1 + AnchoColumna, Y + 1 + (VirtualHeight / 15)
'Escribo los datos.
MyRect.Left = X + 3
MyRect.Right = X + AnchoColumna - 3
MyRect.Top = Y + 1
MyRect.Bottom = Y + (VirtualHeight / 15)
'Color del texto dependiendo si esta la celda seleccionada o no.
If Row = LRow And Col <> LCol And LSelRow = True Then
    PFlex.ForeColor = LSelForeColor
Else
'Color del texto de la celda especificada.
    PFlex.ForeColor = IIf(ArrayCell(3) <> "", ArrayCell(3), LForeColor)
End If
'lSuccess = DrawText(PFlex.hdc, ArrayCell(1), Len(ArrayCell(1)), MyRect, DT_VCENTER Or IIf(ColeccionAlineacion.Item(Col) = "<", DT_LEFT, IIf(ColeccionAlineacion.Item(Col) = ">", DT_RIGHT, DT_CENTER)) Or IIf(LSeeAllText = True, DT_WORDBREAK, DT_SINGLELINE))
Select Case ColeccionAlineacion.item(Col)
Case "<" 'Alineacion a la izquierda.
    lSuccess = DrawText(PFlex.hDC, ArrayCell(1), Len(ArrayCell(1)), MyRect, DT_VCENTER Or DT_LEFT Or IIf(LSeeAllText = True, DT_WORDBREAK, DT_SINGLELINE))
Case ">" 'alineacion a la derecha
    lSuccess = DrawText(PFlex.hDC, ArrayCell(1), Len(ArrayCell(1)), MyRect, DT_VCENTER Or DT_RIGHT Or IIf(LSeeAllText = True, DT_WORDBREAK, DT_SINGLELINE))
Case "^" 'Centrado.
    lSuccess = DrawText(PFlex.hDC, ArrayCell(1), Len(ArrayCell(1)), MyRect, DT_VCENTER Or DT_CENTER Or IIf(LSeeAllText = True, DT_WORDBREAK, DT_SINGLELINE))
End Select

'Dibujo el marco intermitente de seleccion de celda.
If Row = LRow And Col = LCol And FocusRect <> None Then
    Select Case FocusRect
    Case Light
        PFlex.ForeColor = vbBlack
        MyRect.Left = X + 1
        MyRect.Right = X + AnchoColumna
        MyRect.Top = Y + 1
        MyRect.Bottom = Y + (VirtualHeight / 15)
        DrawFocusRect PFlex.hDC, MyRect
    Case Heavy
        PFlex.FillStyle = 1
        PFlex.DrawWidth = 2
        PFlex.ForeColor = LSelBackColor
        Rectangle PFlex.hDC, X + 2, Y + 2, X + AnchoColumna, Y + (VirtualHeight / 15)
        PFlex.DrawWidth = 1
        PFlex.FillStyle = 0
    End Select
End If

PFlex.ScaleMode = vbTwips
'Marco general que envuelve a las celdas.
PFlex.FillStyle = 1
PFlex.Line (0, VirtualHeight)-(AnchoTotal, AltoTotal), vbBlack, B
PFlex.FillStyle = 0
'Fuente del texto de la celda predeterminada _
 por si despues hay cambios.
PFlex.FontName = UserControl.FontName
PFlex.Font.Size = UserControl.FontSize
PFlex.Font.Bold = UserControl.FontBold
End Function

Public Function ResetProperties()
Dim i As Integer, e As Integer
'Elimina el formato de las celdas.
If LRows > 0 Then
    For i = 1 To UBound(ArrayRows)
        ArrayCols = ArrayRows(i)
        For e = 1 To LCols
            ArrayCell = ArrayCols(e)
            ReDim Preserve ArrayCell(1)
            ReDim Preserve ArrayCell(10)
            ArrayCols(e) = ArrayCell
        Next e
        ArrayRows(i) = ArrayCols
    Next i
    Call DibujarCeldas
End If
End Function

Private Function CalcularMouseRowCol(X As Single, Y As Single)
Dim e As Integer
Dim X1 As Variant, X2 As Variant
Dim Y1 As Variant, Y2 As Variant
If Y <= VirtualHeight And X <= AnchoTotal Then
    UserControl.MousePointer = 99
    UserControl.MouseIcon = UserControl.Picture
Else
    UserControl.MousePointer = 0
End If
If LRows > 0 Then
    If Y <= VirtualHeight And Rows > 0 Then
        MouseRow = VsVertical.Value
    Else
        If Y > AltoTotal Then
            MouseRow = VsVertical.Max
        Else
            Y1 = VirtualHeight
            Y2 = VirtualHeight + VirtualHeight
            For e = VsVertical.Value To VsVertical.Max
                If Y > Y1 And Y < Y2 Then
                    MouseRow = e
                    Exit For
                Else
                    Y1 = Y1 + VirtualHeight
                    Y2 = Y1 + VirtualHeight
                End If
            Next e
        End If
    End If
Else
    MouseRow = 0
End If

'Determina MouseCol.
If ColeccionColWidth.Count > 0 Then
    If LCols > 0 And LRows > 0 Then
        X1 = 0
        X2 = ColeccionColWidth(HsHorizontal.Value)
        For e = HsHorizontal.Value To HsHorizontal.Max
            If X >= X1 And X <= X2 Then
                MouseCol = e
                Exit For
            Else
                If X > AnchoTotal And HsHorizontal.Value <> HsHorizontal.Max Then
                    X1 = AnchoTotal - ColeccionColWidth(HsHorizontal.Max)
                    X2 = AnchoTotal
                    MouseCol = HsHorizontal.Max
                    Exit For
                ElseIf X > AnchoTotal And HsHorizontal.Value = HsHorizontal.Max Then
                    X1 = AnchoTotal - ColeccionColWidth(HsHorizontal.Max)
                    X2 = AnchoTotal
                    MouseCol = HsHorizontal.Max
                    Exit For
                Else
                    X1 = X1 + ColeccionColWidth.item(e)

                    If e + 1 <= ColeccionColWidth.Count Then
                        X2 = X1 + ColeccionColWidth.item(e + 1)
                    Else
                        X2 = X1 + ColeccionColWidth.item(e)
                        MouseCol = e
                        Exit For
                    End If
                End If
            End If
        Next e
    Else
        MouseCol = 0
    End If
Else
    If LCols > 0 And LRows > 0 Then
        X1 = 0
        X2 = LColWidth
        For e = HsHorizontal.Value To HsHorizontal.Max
            If X > X1 And X < X2 Then
                MouseCol = e
                Exit For
            Else
                X1 = X1 + LColWidth
                X2 = X1 + LColWidth
            End If
        Next e
    End If
End If
End Function

Public Function GetCellsText(ByVal FirstRow As Integer, ByVal LastRow As Integer, _
        ByVal FirstCol As Integer, ByVal Lastcol As Integer)
Dim i As Integer, e As Integer, CellText As String

For i = FirstRow To LastRow
    For e = FirstCol To Lastcol
        CellText = IIf(e <> Lastcol, CellText & Datos(i, e) & vbTab, IIf(i <> LastRow, CellText & Datos(i, e) & vbCrLf, CellText & Datos(i, e)))
    Next e
Next i
Clipboard.Clear ' Borra el Portapapeles.

Clipboard.SetText CellText    ' Pone texto en el Portapapeles.

End Function


Public Property Get SeeAllText() As Boolean
Attribute SeeAllText.VB_Description = "Devuelve o establece  si se vera el texto de la celda en una linea o no."
SeeAllText = LSeeAllText
End Property

Public Property Let SeeAllText(ByVal vNewValue As Boolean)
LSeeAllText = vNewValue
Call DibujarCeldas
PropertyChanged "SeeAllText"
End Property


Public Function RemoveItem(Optional ByVal Index As Integer)
Dim i As Integer, ArrayConsulta As Variant ', ArrayRows1 As Variant
Select Case Index
Case Is = 0 'Se eliminará la ultima fila.
    Rows = Rows - 1
Case Else
    Select Case Index
    Case Is = 1
        If UBound(ArrayRows) > 1 Then
'Metodo 1
'                        For i = 2 To UBound(ArrayRows)
'                            If i = 2 Then
'                                ReDim ArrayRows1(1)
'                                ArrayRows1(1) = ArrayRows(i)
'                            Else
'                                ReDim Preserve ArrayRows1(UBound(ArrayRows1) + 1)
'                                ArrayRows1(UBound(ArrayRows1)) = ArrayRows(i)
'                            End If
'                        Next i
'                        Set ArrayRows = Nothing
'                        ArrayRows = ArrayRows1
'                        Set ArrayRows1 = Nothing
''''''''''''''''''''''''''''''''''''''''''
'Metodo 2
            For i = 2 To UBound(ArrayRows)
                ArrayRows(i - 1) = ArrayRows(i)
            Next i
            ReDim Preserve ArrayRows(UBound(ArrayRows) - 1)
'No me gusta ninguno de los 2.
            Rows = Rows - 1
        Else
            Rows = Rows - 1
        End If
    Case Is = UBound(ArrayRows) 'Eliminamos la ultima fila.
        Rows = Rows - 1
    Case Else 'eliminamos cualquier fila entre la segunda y la penultima incluidas.
'Metodo 1
'                        ArrayRows1 = ArrayRows
'                        ReDim Preserve ArrayRows1(Index - 1)
'                        a = UBound(ArrayRows1)
'                        ReDim Preserve ArrayRows1(LRows - 1)
'
'                        For i = Index + 1 To UBound(ArrayRows)
'                           ArrayRows1(a + 1) = ArrayRows(i)
'                           a = a + 1
'                        Next i
'''''''''''
'Metodo 2
        ArrayConsulta = ArrayRows(Index) '* deshabilitando estas 2 lineas si el index es mayor que
        Set ArrayConsulta = Nothing      '* ubound(arrayrows) no dara error y se eliminara la ultima fila.
'puesto que la intruccion For i quedara sin efecto y pasaremos a la siguiente instrucccion.

        For i = Index + 1 To UBound(ArrayRows)
            ArrayRows(i - 1) = ArrayRows(i)
        Next i
        ReDim Preserve ArrayRows(UBound(ArrayRows) - 1)
'''''''''''
'Metodo 3
'                        For i = 1 To UBound(ArrayRows)
'                            If i <> Index Then
'                                If i = 1 Then
'                                    ReDim ArrayRows1(1)
'                                    ArrayRows1(1) = ArrayRows(i)
'                                Else
'                                    ReDim Preserve ArrayRows1(UBound(ArrayRows1) + 1)
'                                    ArrayRows1(UBound(ArrayRows1)) = ArrayRows(i)
'                                End If
'                            End If
'                        Next i
'                        Set ArrayRows = Nothing
'                        ArrayRows = ArrayRows1
'                        Set ArrayRows1 = Nothing
'No me gusta ninguno de los 3.
        Rows = Rows - 1
    End Select
End Select
PFlex.Refresh
End Function

Public Property Get ColWidth(ByVal Col As Integer) As Integer
Attribute ColWidth.VB_Description = "Ancho de la columna en Twips."
Attribute ColWidth.VB_MemberFlags = "400"
If ColeccionColWidth.Count > 0 Then
    ColWidth = ColeccionColWidth.item(Col)
Else
    ColWidth = 960
End If
End Property

Public Property Let ColWidth(ByVal Col As Integer, ByVal vNewValue As Integer)
Dim i As Integer
vNewValue = Int(vNewValue / 15) * 15 'multiplos de 15
If ColeccionColWidth.Count = 0 Then
    For i = 1 To LCols
        If i = Col Then
            If vNewValue > 120 Then ColeccionColWidth.Add vNewValue Else ColeccionColWidth.Add 120
        Else
            ColeccionColWidth.Add 960
        End If
    Next i
Else

    If Col = ColeccionColWidth.Count Then
        ColeccionColWidth.Remove Col
        If vNewValue > 120 Then ColeccionColWidth.Add vNewValue Else ColeccionColWidth.Add 120
    Else
        ColeccionColWidth.Remove Col
        If vNewValue > 120 Then ColeccionColWidth.Add vNewValue, , Col Else ColeccionColWidth.Add 120, , Col
    End If
End If

Call CalcularAnchoTotal
Call DibujarCeldas
PFlex.Refresh

End Property

Public Property Get ScrollTrack() As Boolean
Attribute ScrollTrack.VB_Description = "Devuelve o establece si un control MlcGrid debe desplazar su contenido cuando el usuario mueve la caja de desplazamiento por las barras de scroll."
ScrollTrack = LScrollTrack
End Property

Public Property Let ScrollTrack(ByVal vNewValue As Boolean)
LScrollTrack = vNewValue
PropertyChanged "ScrollTrack"
End Property

Private Sub VsVertical_Scroll()
If ScrollTrack = True Then
    Call CalcularAltoTotal
    Call DibujarCeldas
    PFlex.Refresh
End If
End Sub

Public Property Get FocusRect() As TipoFocus
Attribute FocusRect.VB_Description = "Devuelve o establece si se dibujará un rectángulo light o heavy sobre la celda seleccionada."
FocusRect = ValorFocus
End Property

Public Property Let FocusRect(ByVal vNewValue As TipoFocus)
ValorFocus = vNewValue
If Row <> 0 And Col <> 0 Then Call DibujarCelda(Row, Col)
PropertyChanged "FocusRect"
End Property

Public Property Get SelRow() As Boolean
Attribute SelRow.VB_Description = "Devuelve o establece si se marcará toda la fila seleccionada o no."
SelRow = LSelRow
End Property

Public Property Let SelRow(ByVal vNewValue As Boolean)
LSelRow = vNewValue
Call DibujarCeldas
PropertyChanged "SelRow"
End Property

Public Property Get AutoScrolls() As Boolean
Attribute AutoScrolls.VB_Description = "Devuelve o establece si las barras de desplazamiento se harán visibles automáticamente o no."
AutoScrolls = LAutoScrolls
End Property

Public Property Let AutoScrolls(ByVal vNewValue As Boolean)
LAutoScrolls = vNewValue
Call CalcularAltoTotal
Call CalcularAnchoTotal
PropertyChanged "AutoScrolls"
End Property

Public Property Get AllowUserChangeColPos() As Boolean
Attribute AllowUserChangeColPos.VB_Description = "Devuelve o establece si se permite al usuario cambiar las columnas de posición o no."
AllowUserChangeColPos = LAllowUserChangeColPos
End Property

Public Property Let AllowUserChangeColPos(ByVal vNewValue As Boolean)
LAllowUserChangeColPos = vNewValue
PropertyChanged "AllowUserChangeColPos"
End Property

Public Property Get AllowUserSortCol() As Boolean
Attribute AllowUserSortCol.VB_Description = "Devuelve o establece si se permite al usuario ordenarlas columnas o no."
AllowUserSortCol = LAllowUserSortCol
End Property

Public Property Let AllowUserSortCol(ByVal vNewValue As Boolean)
LAllowUserSortCol = vNewValue
PropertyChanged "AllowUserSortCol"
End Property

Public Function ClearRow(ByVal Row As Integer, PreserveText As Boolean)
Dim i As Integer
ArrayCols = ArrayRows(Row)
For i = 1 To LCols
    If PreserveText = True Then ArrayCell = ArrayCols(i): ReDim Preserve ArrayCell(1): _
    ReDim Preserve ArrayCell(10) Else ReDim ArrayCell(10)
    ArrayCols(i) = ArrayCell
Next i
ArrayRows(Row) = ArrayCols
If Row >= VsVertical.Value And Row <= UltimaFila Then
    Call DibujarCeldas
End If
End Function

Public Function ClearCol(ByVal Col As Integer, PreserveText As Boolean)
Dim i As Integer
For i = 1 To Rows
    ArrayCols = ArrayRows(i)
    If PreserveText = True Then ArrayCell = ArrayCols(Col): ReDim Preserve ArrayCell(1): _
    ReDim Preserve ArrayCell(10) Else ReDim ArrayCell(10)
    ArrayCols(Col) = ArrayCell
    ArrayRows(i) = ArrayCols
Next i
Call DibujarCeldas
End Function

Public Function DelGrid()
Cols = 0
End Function

Public Function DelRows()
Rows = 0
End Function
