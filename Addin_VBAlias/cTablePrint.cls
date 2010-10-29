VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTablePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'############################################
'# Grid printing class                      #
'# Author: Jonas Wolz (jwolzvb@yahoo.de)    #
'#------------------------------------------#
'# You are free to use this class in your   #
'# own projects (without paying me a fee).  #
'# If you redistribute this class in        #
'# source form a notification would be      #
'# appreciated.                             #
'# This project comes with absolutely NO    #
'# warranty ! Use it at your own risk !!!   #
'# Please note: I've created the raw        #
'# structure of this class with the class   #
'# builder utility (I didn't want to write  #
'# all those Property Gets and Lets). So    #
'# the German comments inside most of those #
'# Property Get/Lets (and near private      #
'# variables) are only the utility's        #
'# standard comments. (Don't get confused   #
'# by them)                                 #
'############################################

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const c_DTDefFmt = DT_NOPREFIX 'Or DT_SINGLELINE Or DT_VCENTER


Public Enum EAlignment
    eRight = DT_RIGHT
    eLeft = DT_LEFT
    eCenter = DT_CENTER
End Enum

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarCols As Long 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarRows As Long 'lokale Kopie
Private mvarTextMatrix() As String 'lokale Kopie
'Private mvarColFont() As StdFont 'lokale Kopie
Private mvarHeaderText() As String 'lokale Kopie
Private mvarHeaderFont() As StdFont 'lokale Kopie
Private mvarColAlignment() As EAlignment 'lokale Kopie
Private mvarHasFooter As Boolean 'lokale Kopie
Private mvarFooterFont() As StdFont 'lokale Kopie
Private mvarFooterText() As String 'lokale Kopie
Private mvarMarginTop As Single 'lokale Kopie
Private mvarMarginBottom As Single 'lokale Kopie
Private mvarMarginLeft As Single 'lokale Kopie
'Private mvarMarginRight As Single 'lokale Kopie
Private mvarLineThickness As Integer 'lokale Kopie
Private mvarHeaderLineThickness As Integer 'lokale Kopie
Private mvarFooterLineThickness As Integer 'lokale Kopie
Private mvarCellYOffset As Single 'lokale Kopie
Private mvarCellXOffset As Single 'lokale Kopie
Private m_ColWidth() As Single
Private m_MergeCol() As Boolean
Private m_FontMatrix() As StdFont
Private mvarHeaderRows As Long 'lokale Kopie
Private mvarPictureMatrix() As IPictureDisp 'lokale Kopie
Private m_MergeHeaderCol() As Boolean
Private m_MergeHeaderRow() As Boolean
Private m_MergeRow() As Boolean

'Um dieses Ereignis auszulösen, verwenden Sie RaiseEvent mit der folgenden Syntax:
'RaiseEvent NewPage[(arg1, arg2, ... , argn)]
Public Event NewPage(objOutput As Object, TopMarginAlreadySet As Boolean, bCancel As Boolean, ByVal lLastPrintedRow As Long)

Public PrintHeaderOnEveryPage As Boolean
Public CenterMergedHeader As Boolean
Public RowHeightMin As Single, HeaderRowHeightMin As Single, FooterRowHeightMin As Single
Public ResizeCellsToPicHeight As Boolean

Function CalcNumRowsPerPage(objOutput As Object, Optional ByVal bWithHeader = True, Optional ByVal bWithFooter)
Dim sngNormalRowH As Single, sngHeaderRowH As Single, sngFooterRowH As Single
Dim LastFont As StdFont, sngTmp As Single, LRow As Long, L As Long

If IsMissing(bWithFooter) Then
    bWithFooter = mvarHasFooter
End If

sngNormalRowH = RowHeightMin - 2 * mvarCellYOffset
sngHeaderRowH = HeaderRowHeightMin - 2 * mvarCellYOffset
sngFooterRowH = FooterRowHeightMin - 2 * mvarCellYOffset

For L = 0 To mvarCols - 1
    For LRow = 0 To mvarRows - 1
        If Not (LastFont Is m_FontMatrix(LRow, L)) Then
            Set LastFont = m_FontMatrix(LRow, L)
            Set objOutput.Font = LastFont
        End If
        sngTmp = objOutput.TextHeight(mvarTextMatrix(LRow, L))
        If sngTmp > sngNormalRowH Then sngNormalRowH = sngTmp
        If ResizeCellsToPicHeight Then
            If Not (mvarPictureMatrix(LRow, L) Is Nothing) Then
                sngTmp = objOutput.ScaleY(mvarPictureMatrix(LRow, L).Height, vbHimetric, objOutput.ScaleMode)
                If sngTmp > sngNormalRowH Then sngNormalRowH = sngTmp
            End If
        End If
    Next
    For LRow = 0 To mvarHeaderRows - 1
        If Not (LastFont Is mvarHeaderFont(LRow, L)) Then
            Set LastFont = mvarHeaderFont(LRow, L)
            Set objOutput.Font = LastFont
        End If
        sngTmp = objOutput.TextHeight(mvarHeaderText(LRow, L))
        If sngTmp > sngHeaderRowH Then sngHeaderRowH = sngTmp
    Next
    If mvarHasFooter Then
        If Not (LastFont Is mvarFooterFont(L)) Then
            Set LastFont = mvarFooterFont(L)
            Set objOutput.Font = LastFont
        End If
        sngTmp = objOutput.TextHeight(mvarFooterText(L))
        If sngTmp > sngFooterRowH Then sngFooterRowH = sngTmp
    End If
Next

sngNormalRowH = sngNormalRowH + 2 * mvarCellYOffset
sngHeaderRowH = sngHeaderRowH + 2 * mvarCellYOffset
If sngFooterRowH > 0 Then sngFooterRowH = sngFooterRowH + 2 * mvarCellYOffset

sngTmp = (objOutput.ScaleHeight - mvarMarginBottom - mvarMarginTop)
'CalcNumRowsPerPage = Int((objOutput.ScaleHeight - sngHeaderRowH - sngFooterRowH - mvarMarginBottom - mvarMarginTop) / sngNormalRowH)
If bWithHeader Then
'CalcNumRowsPerPage = CalcNumRowsPerPage + 1
    sngTmp = sngTmp - sngHeaderRowH * mvarHeaderRows
End If
If bWithFooter Then
'CalcNumRowsPerPage = CalcNumRowsPerPage + 1
    sngTmp = sngTmp - sngFooterRowH
End If
CalcNumRowsPerPage = Int(sngTmp / sngNormalRowH)
End Function

Private Sub pDoVCenter(ByVal hDC As Long, rcDraw As RECT, sText As String)
Dim rcTmp As RECT, lRet As Long
lRet = DrawText(hDC, sText, -1, rcTmp, c_DTDefFmt Or DT_CALCRECT)
rcDraw.Top = (rcDraw.Bottom - rcDraw.Top - lRet) \ 2 + rcDraw.Top
End Sub


Public Property Set PictureMatrix(ByVal Row As Long, ByVal Col As Long, ByVal vData As IPictureDisp)
Dim L As Long, L2 As Long
If (Row < 0) And (Col < 0) Then
    For L = 0 To mvarRows - 1
        For L2 = 0 To mvarCols - 1
            Set mvarPictureMatrix(L, L2) = vData
        Next
    Next
ElseIf (Row < 0) Then
    For L = 0 To mvarRows - 1
        Set mvarPictureMatrix(L, Col) = vData
    Next
ElseIf (Col < 0) Then
    For L = 0 To mvarCols - 1
        Set mvarPictureMatrix(Row, L) = vData
    Next
Else
    Set mvarPictureMatrix(Row, Col) = vData
End If

End Property


Public Property Get PictureMatrix(ByVal Row As Long, ByVal Col As Long) As IPictureDisp
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.PictureMatrix
Set PictureMatrix = mvarPictureMatrix(Row, Col)
End Property



Public Property Let HeaderRows(ByVal vData As Long)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.HeaderRows = 5
If vData < 1 Then
    Err.Raise 380
End If
mvarHeaderRows = vData
pRedimArrays
End Property


Public Property Get HeaderRows() As Long
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.HeaderRows
HeaderRows = mvarHeaderRows
End Property



Public Sub DrawTable(objOutput As Object, Optional lRowToStart As Long = 0, Optional lStartPage As Long = 1)
Dim rctDraw As RECT, LastRow As Long, PgNum As Long
Dim sngNormalRowH As Single, sngHeaderRowH As Single, sngFooterRowH As Single
Dim sngSW As Single, L As Long, sngTmp As Single, bFlag As Boolean
Dim sngYStart As Single, iSM As Integer, lArrCW() As Long, lMargPix As Long
Dim LRow As Long, lRH As Long, StartRow As Long, LastTop As Single
Dim lPixXOffset As Long, OldFont As StdFont, bMerged() As Boolean, LastFont As StdFont
Dim picCoords(1 To 4) As Single, lPixYOffset As Long, lTmp As Long
Dim lHeaderMergeRow() As Long, bHeaderFlag() As Boolean

'sngSW = objOutput.ScaleWidth - mvarMarginRight

Set OldFont = objOutput.Font
Set LastFont = OldFont
ReDim lArrCW(0 To mvarCols - 1)
ReDim lHeaderMergeRow(0 To mvarCols - 1) ' Below again (Loop)
ReDim bHeaderFlag(0 To mvarHeaderRows - 1)
ReDim bMerged(0 To mvarHeaderRows - 1)

iSM = objOutput.ScaleMode
sngSW = mvarMarginLeft

sngNormalRowH = RowHeightMin - 2 * mvarCellYOffset
sngHeaderRowH = HeaderRowHeightMin - 2 * mvarCellYOffset
sngFooterRowH = FooterRowHeightMin - 2 * mvarCellYOffset

For L = 0 To mvarCols - 1
    For LRow = 0 To mvarRows - 1
        If Not (LastFont Is m_FontMatrix(LRow, L)) Then
            Set LastFont = m_FontMatrix(LRow, L)
            Set objOutput.Font = LastFont
        End If
        sngTmp = objOutput.TextHeight(mvarTextMatrix(LRow, L))
        If sngTmp > sngNormalRowH Then sngNormalRowH = sngTmp
        If ResizeCellsToPicHeight Then
            If Not (mvarPictureMatrix(LRow, L) Is Nothing) Then
                sngTmp = objOutput.ScaleY(mvarPictureMatrix(LRow, L).Height, vbHimetric, iSM)
                If sngTmp > sngNormalRowH Then sngNormalRowH = sngTmp
            End If
        End If
    Next
    For LRow = 0 To mvarHeaderRows - 1
        If Not (LastFont Is mvarHeaderFont(LRow, L)) Then
            Set LastFont = mvarHeaderFont(LRow, L)
            Set objOutput.Font = LastFont
        End If
        sngTmp = objOutput.TextHeight(mvarHeaderText(LRow, L))
        If sngTmp > sngHeaderRowH Then sngHeaderRowH = sngTmp
    Next
    If mvarHasFooter Then
        If Not (LastFont Is mvarFooterFont(L)) Then
            Set LastFont = mvarFooterFont(L)
            Set objOutput.Font = LastFont
        End If
        sngTmp = objOutput.TextHeight(mvarFooterText(L))
        If sngTmp > sngFooterRowH Then sngFooterRowH = sngTmp
    End If
    sngSW = sngSW + m_ColWidth(L)
    lArrCW(L) = objOutput.ScaleX(m_ColWidth(L), iSM, vbPixels)
    lHeaderMergeRow(L) = -1
Next
lMargPix = objOutput.ScaleX(mvarMarginLeft, iSM, vbPixels)
lPixXOffset = objOutput.ScaleX(mvarCellXOffset, iSM, vbPixels)
lPixYOffset = objOutput.ScaleY(mvarCellYOffset, iSM, vbPixels)

sngNormalRowH = sngNormalRowH + 2 * mvarCellYOffset
sngHeaderRowH = sngHeaderRowH + 2 * mvarCellYOffset
If sngFooterRowH > 0 Then sngFooterRowH = sngFooterRowH + 2 * mvarCellYOffset

'objOutput.CurrentX = mvarMarginLeft
objOutput.FillStyle = vbFSTransparent
PgNum = lStartPage
LastRow = lRowToStart - 1
Do
    sngTmp = 0
    sngTmp = (objOutput.ScaleHeight - objOutput.CurrentY - mvarMarginBottom)
'L = Int((objOutput.ScaleHeight - objOutput.CurrentY - sngHeaderRowH - sngFooterRowH - mvarMarginBottom) / sngNormalRowH)
    If sngTmp >= sngNormalRowH Then

        If LastRow >= mvarRows Then
'LastRow = LastRow + 1 'We don't need the footer !
            sngTmp = sngTmp - sngFooterRowH
        End If
        If PrintHeaderOnEveryPage Or (PgNum > 1) Then
'LastRow = LastRow + 1
            sngTmp = sngTmp - sngHeaderRowH * mvarHeaderRows
        End If
        StartRow = LastRow + 1
        L = Int(sngTmp / sngNormalRowH)
        LastRow = LastRow + L
        If LastRow > mvarRows - 1 Then
            LastRow = mvarRows - 1
        End If

        sngYStart = mvarMarginTop 'objOutput.CurrentY
        LastTop = sngYStart
        If PrintHeaderOnEveryPage Or (PgNum = 1) Then
            objOutput.DrawWidth = mvarHeaderLineThickness
            For LRow = 0 To mvarHeaderRows - 1
'                    rctDraw.Bottom = rctDraw.Bottom + objOutput.ScaleY(sngHeaderRowH, iSM, vbPixels)
'objOutput.Line (mvarMarginLeft, sngYStart + lRow * sngHeaderRowH)-(sngSW, sngYStart + sngHeaderRowH * (lRow + 1)), , B
'objOutput.Line (mvarMarginLeft, sngYStart + sngHeaderRowH)-(sngSW, sngYStart )
'                    rctDraw.Top = objOutput.ScaleY(sngYStart + lRow * sngHeaderRowH, iSM, vbPixels)
                rctDraw.Left = lMargPix + lPixXOffset
                rctDraw.Right = lMargPix - lPixXOffset
                sngTmp = mvarMarginLeft
                bHeaderFlag(LRow) = False
                bMerged(LRow) = False
                For L = 0 To mvarCols - 1
                    If m_MergeHeaderCol(L) Then
                        If LRow < mvarHeaderRows - 1 Then
                            If mvarHeaderText(LRow, L) = mvarHeaderText(LRow + 1, L) Then
                                bFlag = False
                                If lHeaderMergeRow(L) = -1 Then lHeaderMergeRow(L) = LRow
                            Else
                                bFlag = True
                            End If
                        Else
                            bFlag = True
                        End If
                    Else
                        lHeaderMergeRow(L) = -1
                        bFlag = True
                    End If
                    rctDraw.Right = rctDraw.Right + lArrCW(L) '- 2 * lPixXOffset
                    If bFlag Then
                        If lHeaderMergeRow(L) = -1 Then
                            rctDraw.Top = objOutput.ScaleY(sngYStart + LRow * sngHeaderRowH, iSM, vbPixels)
                            rctDraw.Bottom = rctDraw.Top + objOutput.ScaleY(sngHeaderRowH, iSM, vbPixels)
                        Else
                            rctDraw.Top = objOutput.ScaleY(sngYStart + lHeaderMergeRow(L) * sngHeaderRowH, iSM, vbPixels)
                            rctDraw.Bottom = rctDraw.Top + objOutput.ScaleY((LRow - lHeaderMergeRow(L) + 1) * sngHeaderRowH, iSM, vbPixels)
                            lHeaderMergeRow(LRow) = -1
                        End If
                        If Not (LastFont Is mvarHeaderFont(LRow, L)) Then
                            Set LastFont = mvarHeaderFont(LRow, L)
                            Set objOutput.Font = LastFont
                        End If
                        If m_MergeHeaderRow(LRow) Then
                            If L < mvarCols - 1 Then
                                If mvarHeaderText(LRow, L) = mvarHeaderText(LRow, L + 1) Then
                                    bHeaderFlag(LRow) = False
                                    bMerged(LRow) = True
                                Else
                                    If bMerged(LRow) And bHeaderFlag(LRow) Then bMerged(LRow) = False
                                    bHeaderFlag(LRow) = True
                                End If
                            Else
                                bHeaderFlag(LRow) = True
                            End If
                        Else
                            bHeaderFlag(LRow) = True
                        End If
                        If bHeaderFlag(LRow) Then
                            objOutput.ScaleMode = vbPixels
                            objOutput.Line (rctDraw.Left - lPixXOffset, rctDraw.Top)-(rctDraw.Right + lPixXOffset, rctDraw.Bottom), , B
                            objOutput.ScaleMode = iSM

                            pDoVCenter objOutput.hDC, rctDraw, mvarHeaderText(LRow, L)
                            DrawText objOutput.hDC, mvarHeaderText(LRow, L), -1, rctDraw, c_DTDefFmt Or IIf((bMerged(LRow) And CenterMergedHeader), DT_CENTER, mvarColAlignment(L))
                        End If
                    End If
                    sngTmp = sngTmp + m_ColWidth(L)
                    If bHeaderFlag(LRow) Then rctDraw.Left = rctDraw.Right + 2 * lPixXOffset
                Next
            Next
            LastTop = LastTop + sngHeaderRowH * mvarHeaderRows
        End If
        lRH = objOutput.ScaleY(sngNormalRowH, iSM, vbPixels)
        objOutput.DrawWidth = mvarLineThickness
        ReDim lHeaderMergeRow(0 To mvarRows - 1)
        For L = 0 To UBound(lHeaderMergeRow())
            lHeaderMergeRow(L) = -1
'                bHeaderFlag(L) = False
        Next
        sngYStart = LastTop
        rctDraw.Left = lMargPix + lPixXOffset
        sngTmp = mvarMarginLeft
        For L = 0 To mvarCols - 1
'                If Not (objOutput.Font Is mvarColFont(L)) Then
'                    Set objOutput.Font = mvarColFont(L)
'                End If
            rctDraw.Top = objOutput.ScaleY(sngYStart, iSM, vbPixels)
            LastTop = sngYStart
            rctDraw.Right = rctDraw.Left + lArrCW(L) - 2 * lPixXOffset
            rctDraw.Bottom = rctDraw.Top
            bFlag = True
            For LRow = StartRow To LastRow
                If Not (LastFont Is m_FontMatrix(LRow, L)) Then
                    Set LastFont = m_FontMatrix(LRow, L)
                    Set objOutput.Font = LastFont
                End If
                rctDraw.Bottom = rctDraw.Bottom + lRH
'                    If bFlag Then objOutput.Line (sngTmp, LastTop)-(sngTmp + m_ColWidth(L), LastTop)
                If m_MergeCol(L) Then
                    If LRow < LastRow Then
                        If mvarTextMatrix(LRow, L) = mvarTextMatrix(LRow + 1, L) Then
                            bFlag = False
                        Else
                            bFlag = True
                        End If
                    Else
                        bFlag = True
                    End If
                Else
                    bFlag = True
                End If
                If bFlag Then
                    If m_MergeRow(LRow) Then
                        If L < mvarCols - 1 Then
                            If mvarTextMatrix(LRow, L) = mvarTextMatrix(LRow, L + 1) Then
'Now ColPos:
                                If lHeaderMergeRow(LRow) = -1 Then lHeaderMergeRow(LRow) = rctDraw.Left 'L
                                bMerged(0) = False
                            Else
                                bMerged(0) = True
                            End If
                        Else
                            bMerged(0) = True
                        End If
                    Else
                        bMerged(0) = True
                        lHeaderMergeRow(LRow) = -1
                    End If

                    If bMerged(0) Then
                        If lHeaderMergeRow(LRow) <> -1 Then
                            lTmp = rctDraw.Left
                            rctDraw.Left = lHeaderMergeRow(LRow)
                        End If
                        If Not (mvarPictureMatrix(LRow, L) Is Nothing) Then
                            picCoords(3) = objOutput.ScaleX(mvarPictureMatrix(LRow, L).Width, vbHimetric, iSM)
                            picCoords(4) = objOutput.ScaleY(mvarPictureMatrix(LRow, L).Height, vbHimetric, iSM)
                            If picCoords(4) > (m_ColWidth(L) - 2 * mvarCellXOffset) Then picCoords(4) = (m_ColWidth(L) - 2 * mvarCellXOffset)
                            picCoords(2) = (objOutput.ScaleY(rctDraw.Bottom - rctDraw.Top, vbPixels, iSM) - picCoords(4)) / 2
                            If picCoords(2) < 0 Then
                                picCoords(3) = picCoords(3) + 2 * picCoords(2)
                                picCoords(2) = 0
                            End If
                            picCoords(2) = picCoords(2) + LastTop
                            If mvarColAlignment(L) = eLeft Then
                                picCoords(1) = objOutput.ScaleX(rctDraw.Left, vbPixels, iSM)
                            ElseIf mvarColAlignment(L) = eRight Then
                                picCoords(1) = sngTmp + (m_ColWidth(L) - mvarCellXOffset - picCoords(4))
                            ElseIf mvarColAlignment(L) = eCenter Then
                                picCoords(1) = sngTmp + (m_ColWidth(L) - picCoords(4)) / 2
                            End If
                            On Error Resume Next
                            objOutput.PaintPicture mvarPictureMatrix(LRow, L), picCoords(1), picCoords(2), , , , , picCoords(3), picCoords(4)
                            On Error GoTo 0
                        End If

                        objOutput.ScaleMode = vbPixels
                        objOutput.Line (rctDraw.Left - lPixXOffset, rctDraw.Top)-(rctDraw.Right + lPixXOffset, rctDraw.Bottom), , B
                        objOutput.ScaleMode = iSM
                        pDoVCenter objOutput.hDC, rctDraw, mvarTextMatrix(LRow, L)
                        DrawText objOutput.hDC, mvarTextMatrix(LRow, L), -1, rctDraw, c_DTDefFmt Or mvarColAlignment(L)
                        If lHeaderMergeRow(LRow) <> -1 Then
                            lHeaderMergeRow(LRow) = -1
                            rctDraw.Left = lTmp
                        End If
                    End If
                    rctDraw.Top = rctDraw.Bottom
                End If
                LastTop = LastTop + sngNormalRowH
            Next
'                objOutput.Line (sngTmp, sngYStart)-(sngTmp, LastTop)
            rctDraw.Left = rctDraw.Right + 2 * lPixXOffset
            sngTmp = sngTmp + m_ColWidth(L)
        Next
'objOutput.Line (sngSW, sngYStart)-(sngSW, LastTop)
        If mvarHasFooter And (LastRow >= mvarRows - 1) Then
            objOutput.DrawWidth = mvarFooterLineThickness
            objOutput.Line (mvarMarginLeft, LastTop)-(sngSW, LastTop + sngFooterRowH), , B
'objOutput.Line (mvarMarginLeft, sngYStart + sngHeaderRowH)-(sngSW, sngYStart )
            lTmp = rctDraw.Top
            rctDraw.Bottom = rctDraw.Top + objOutput.ScaleY(sngHeaderRowH, iSM, vbPixels)
            rctDraw.Left = lMargPix + lPixXOffset
            sngTmp = mvarMarginLeft
            For L = 0 To mvarCols - 1
                If Not (LastFont Is mvarFooterFont(L)) Then
                    Set LastFont = mvarFooterFont(L)
                    Set objOutput.Font = LastFont
                End If
                rctDraw.Right = rctDraw.Left + lArrCW(L) - 2 * lPixXOffset
                rctDraw.Top = lTmp

                pDoVCenter objOutput.hDC, rctDraw, mvarFooterText(L)
                DrawText objOutput.hDC, mvarFooterText(L), -1, rctDraw, c_DTDefFmt Or mvarColAlignment(L)
                rctDraw.Left = rctDraw.Right + 2 * lPixXOffset

                If L > 0 Then  'We don't need to draw the first line
                    objOutput.Line (sngTmp, LastTop)-(sngTmp, LastTop + sngFooterRowH)
                End If
                sngTmp = sngTmp + m_ColWidth(L)
            Next
            LastTop = LastTop + sngFooterRowH
'            Else
'                objOutput.Line (mvarMarginLeft, LastTop)-(sngSW, LastTop)
        End If
    End If
    objOutput.CurrentY = LastTop
    If LastRow < mvarRows - 1 Then
        bFlag = False
        bMerged(0) = False
        RaiseEvent NewPage(objOutput, bFlag, bMerged(0), LastRow)
'bMerged(0) is the cancel flag here !
        If bMerged(0) Then Exit Do
        If Not bFlag Then
            objOutput.CurrentY = mvarMarginTop
        End If
        PgNum = PgNum + 1
    Else
        Exit Do
    End If
    ReDim lHeaderMergeRow(0 To mvarCols - 1)
Loop
Set objOutput.Font = OldFont
End Sub

Public Property Let CellXOffset(ByVal vData As Single)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.CellXOffset = 5
mvarCellXOffset = vData
End Property


Public Property Get CellXOffset() As Single
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.CellXOffset
CellXOffset = mvarCellXOffset
End Property



Public Property Let CellYOffset(ByVal vData As Single)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.CellYOffset = 5
mvarCellYOffset = vData
End Property


Public Property Get CellYOffset() As Single
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.CellYOffset
CellYOffset = mvarCellYOffset
End Property



Public Property Let FooterLineThickness(ByVal vData As Integer)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.FooterLineThickness = 5
mvarFooterLineThickness = vData
End Property


Public Property Get FooterLineThickness() As Integer
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.FooterLineThickness
FooterLineThickness = mvarFooterLineThickness
End Property



Public Property Let HeaderLineThickness(ByVal vData As Integer)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.HeaderLineThickness = 5
mvarHeaderLineThickness = vData
End Property


Public Property Get HeaderLineThickness() As Integer
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.HeaderLineThickness
HeaderLineThickness = mvarHeaderLineThickness
End Property



Public Property Let LineThickness(ByVal vData As Integer)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.LineThickness = 5
mvarLineThickness = vData
End Property


Public Property Get LineThickness() As Integer
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.LineThickness
LineThickness = mvarLineThickness
End Property



'Public Property Let MarginRight(ByVal vData As Single)
''wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
''Syntax: X.MarginRight = 5
'    mvarMarginRight = vData
'End Property


'Public Property Get MarginRight() As Single
''wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
''Syntax: Debug.Print X.MarginRight
'    MarginRight = mvarMarginRight
'End Property



Public Property Let MarginLeft(ByVal vData As Single)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.MarginLeft = 5
mvarMarginLeft = vData
End Property


Public Property Get MarginLeft() As Single
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.MarginLeft
MarginLeft = mvarMarginLeft
End Property



Public Property Let MarginBottom(ByVal vData As Single)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.MarginBottom = 5
mvarMarginBottom = vData
End Property


Public Property Get MarginBottom() As Single
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.MarginBottom
MarginBottom = mvarMarginBottom
End Property



Public Property Let MarginTop(ByVal vData As Single)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.MarginTop = 5
mvarMarginTop = vData
End Property


Public Property Get MarginTop() As Single
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.MarginTop
MarginTop = mvarMarginTop
End Property



Public Property Let FooterText(ByVal Col As Long, ByVal vData As String)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.FooterText = 5
If Col < 0 Then
    pLetAllArrItems mvarFooterText, vData
Else
    mvarFooterText(Col) = vData
End If
End Property


Public Property Get FooterText(ByVal Col As Long) As String
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.FooterText
FooterText = mvarFooterText(Col)
End Property



Public Property Set FooterFont(ByVal Col As Long, ByVal vData As StdFont)
'wird beim Zuweisen eines Objekts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: Set x.FooterFont = Form1
If Col < 0 Then
    pSetAllArrItems mvarFooterFont, vData
Else
    Set mvarFooterFont(Col) = vData
End If
End Property


Public Property Get FooterFont(ByVal Col As Long) As StdFont
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.FooterFont
Set FooterFont = mvarFooterFont(Col)
End Property



Public Property Let HasFooter(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.HasFooter = 5
mvarHasFooter = vData
End Property


Public Property Get HasFooter() As Boolean
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.HasFooter
HasFooter = mvarHasFooter
End Property



Public Property Let ColAlignment(ByVal Col As Long, ByVal vData As EAlignment)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.ColAlignment = 5
If Col < 0 Then
    pLetAllArrItems mvarColAlignment, vData
Else
    mvarColAlignment(Col) = vData
End If
End Property


Public Property Get ColAlignment(ByVal Col As Long) As EAlignment
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.ColAlignment
ColAlignment = mvarColAlignment(Col)
End Property



Public Property Set HeaderFont(ByVal Row As Long, ByVal Col As Long, ByVal vData As Object)
'wird beim Zuweisen eines Objekts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: Set x.HeaderFont = Form1
Dim L As Long, L2 As Long
If (Row < 0) And (Col < 0) Then
    For L = 0 To mvarHeaderRows - 1
        For L2 = 0 To mvarCols - 1
            Set mvarHeaderFont(L, L2) = vData
        Next
    Next
ElseIf (Row < 0) Then
    For L = 0 To mvarHeaderRows - 1
        Set mvarHeaderFont(L, Col) = vData
    Next
ElseIf (Col < 0) Then
    For L = 0 To mvarCols - 1
        Set mvarHeaderFont(Row, L) = vData
    Next
Else
    Set mvarHeaderFont(Row, Col) = vData
End If
End Property


Public Property Get HeaderFont(ByVal Row As Long, ByVal Col As Long) As StdFont
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.HeaderFont
Set HeaderFont = mvarHeaderFont(Row, Col)
End Property



Public Property Let HeaderText(ByVal Row As Long, ByVal Col As Long, ByVal vData As String)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.HeaderText = 5
Dim L As Long, L2 As Long
If (Row < 0) And (Col < 0) Then
    For L = 0 To mvarHeaderRows - 1
        For L2 = 0 To mvarCols - 1
            mvarHeaderText(L, L2) = vData
        Next
    Next
ElseIf (Row < 0) Then
    For L = 0 To mvarHeaderRows - 1
        mvarHeaderText(L, Col) = vData
    Next
ElseIf (Col < 0) Then
    For L = 0 To mvarCols - 1
        mvarHeaderText(Row, L) = vData
    Next
Else
    mvarHeaderText(Row, Col) = vData
End If

End Property


Public Property Get HeaderText(ByVal Row As Long, ByVal Col As Long) As String
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.HeaderText
HeaderText = mvarHeaderText(Row, Col)
End Property



Public Property Set ColFont(ByVal Col As Long, ByVal vData As StdFont)
'    If Col < 0 Then
'        pSetAllArrItems mvarColFont(), vData
'    Else
'        Set mvarColFont(Col) = vData
'    End If
Set FontMatrix(-1, Col) = vData
End Property


Public Property Get ColFont(ByVal Col As Long) As StdFont
'    Set ColFont(Col) = mvarColFont(Col)
Set ColFont = FontMatrix(0, Col)
End Property
Public Property Get MergeHeaderRow(ByVal Row As Long) As Boolean
MergeHeaderRow = m_MergeHeaderRow(Row)
End Property
Public Property Let MergeHeaderRow(ByVal Row As Long, ByVal NewV As Boolean)
m_MergeHeaderRow(Row) = NewV
End Property
Public Property Get MergeHeaderCol(ByVal Col As Long) As Boolean
MergeHeaderCol = m_MergeHeaderCol(Col)
End Property
Public Property Let MergeHeaderCol(ByVal Col As Long, ByVal NewV As Boolean)
m_MergeHeaderCol(Col) = NewV
End Property
Public Property Get MergeRow(ByVal Row As Long) As Boolean
MergeRow = m_MergeRow(Row)
End Property
Public Property Let MergeRow(ByVal Row As Long, ByVal NewV As Boolean)
m_MergeRow(Row) = NewV
End Property
Public Property Get MergeCol(ByVal Col As Long) As Boolean
MergeCol = m_MergeCol(Col)
End Property
Public Property Let MergeCol(ByVal Col As Long, ByVal NewV As Boolean)
m_MergeCol(Col) = NewV
End Property
Private Sub pLetAllArrItems(vArr As Variant, vValue As Variant)
Dim L As Long

For L = LBound(vArr) To UBound(vArr)
    vArr(L) = vValue
Next
End Sub
Private Sub pSetAllArrItems(vArr As Variant, objValue As Object)
Dim L As Long

For L = LBound(vArr) To UBound(vArr)
    Set vArr(L) = objValue
Next
End Sub
Private Sub pRedimArrays()
If mvarRows > 0 Then
    ReDim mvarTextMatrix(0 To mvarRows - 1, 0 To mvarCols - 1)
    ReDim m_FontMatrix(0 To mvarRows - 1, 0 To mvarCols - 1)
    ReDim mvarPictureMatrix(0 To mvarRows - 1, 0 To mvarCols - 1)
Else
    Erase mvarTextMatrix
    Erase m_FontMatrix
    Erase mvarPictureMatrix
End If
'ReDim mvarColFont(0 To mvarCols - 1)
ReDim mvarColAlignment(0 To mvarCols - 1)
ReDim mvarHeaderText(0 To mvarHeaderRows - 1, 0 To mvarCols - 1)
ReDim mvarHeaderFont(0 To mvarHeaderRows - 1, 0 To mvarCols - 1)
ReDim m_ColWidth(0 To mvarCols - 1)
ReDim m_MergeCol(0 To mvarCols - 1)
ReDim m_MergeHeaderCol(0 To mvarCols - 1)
ReDim m_MergeRow(0 To mvarRows - 1)
ReDim m_MergeHeaderRow(0 To mvarHeaderRows - 1)

If mvarHasFooter Then
    ReDim mvarFooterText(0 To mvarCols - 1)
    ReDim mvarFooterFont(0 To mvarCols - 1)
Else
    Erase mvarFooterText
    Erase mvarFooterFont
End If

pLetAllArrItems m_ColWidth, 15
End Sub

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, ByVal vData As String)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.TextMatrix = 5
Dim L As Long, L2 As Long
If (Row < 0) And (Col < 0) Then
    For L = 0 To mvarRows - 1
        For L2 = 0 To mvarCols - 1
            mvarTextMatrix(L, L2) = vData
        Next
    Next
ElseIf (Row < 0) Then
    For L = 0 To mvarRows - 1
        mvarTextMatrix(L, Col) = vData
    Next
ElseIf (Col < 0) Then
    For L = 0 To mvarCols - 1
        mvarTextMatrix(Row, L) = vData
    Next
Else
    mvarTextMatrix(Row, Col) = vData
End If
End Property


Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.TextMatrix
TextMatrix = mvarTextMatrix(Row, Col)
End Property

Public Property Set FontMatrix(ByVal Row As Long, ByVal Col As Long, ByVal vData As StdFont)
Dim L As Long, L2 As Long
If (Row < 0) And (Col < 0) Then
    For L = 0 To mvarRows - 1
        For L2 = 0 To mvarCols - 1
            Set m_FontMatrix(L, L2) = vData
        Next
    Next
ElseIf (Row < 0) Then
    For L = 0 To mvarRows - 1
        Set m_FontMatrix(L, Col) = vData
    Next
ElseIf (Col < 0) Then
    For L = 0 To mvarCols - 1
        Set m_FontMatrix(Row, L) = vData
    Next
Else
    Set m_FontMatrix(Row, Col) = vData
End If
End Property


Public Property Get FontMatrix(ByVal Row As Long, ByVal Col As Long) As StdFont
Set FontMatrix = m_FontMatrix(Row, Col)
End Property

Public Property Let Rows(ByVal vData As Long)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.Rows = 5
If vData < 0 Then
    Err.Raise 380
End If
mvarRows = vData
pRedimArrays
End Property


Public Property Get Rows() As Long
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.Rows
Rows = mvarRows
End Property




Public Property Let Cols(ByVal vData As Long)
'wird beim Zuweisen eines Werts in eine Eigenschaft auf der linken Seite der Gleichung, verwendet.
'Syntax: X.Cols = 5
If vData < 1 Then
    Err.Raise 380
End If
mvarCols = vData
pRedimArrays
End Property


Public Property Get Cols() As Long
'wird beim Ermitteln einer Eignschaft auf der rechten Seite der Gleichung verwendet.
'Syntax: Debug.Print X.Cols
Cols = mvarCols
End Property



Private Sub Class_Initialize()
mvarRows = 5
mvarCols = 2
mvarHasFooter = True
mvarLineThickness = 1
mvarHeaderLineThickness = 1
mvarFooterLineThickness = 1
mvarHeaderRows = 1
pRedimArrays
End Sub



Public Property Get ColWidth(ByVal Col As Long) As Single
ColWidth = m_ColWidth(Col)
End Property

Public Property Let ColWidth(ByVal Col As Long, ByVal vNewValue As Single)
If Col < 0 Then
    pLetAllArrItems m_ColWidth, vNewValue
Else
    m_ColWidth(Col) = vNewValue
End If
End Property

