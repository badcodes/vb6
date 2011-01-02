Attribute VB_Name = "MFilter"
Option Explicit

Public Enum EErrorFilter
    eeBaseFilter = 13490    ' Filter
End Enum

Sub FilterTextFile(filter As IFilter)
    
    BugAssert filter.Source <> sEmpty
    ' Target can be another file or replacement of current file
    Dim sTarget As String, fReplace As Boolean
    sTarget = filter.Target
    If sTarget = sEmpty Or sTarget = filter.Source Then
        sTarget = MUtility.GetTempFile("FLT", ".")
        fReplace = True
    End If
    
    ' Open input file
    On Error GoTo FilterTextError1
    Dim nIn As Integer, nOut As Integer
    nIn = FreeFile
    Open filter.Source For Input Access Read Lock Write As #nIn

    ' Open target output file
    On Error GoTo FilterTextError2
    nOut = FreeFile
    Open sTarget For Output Access Write Lock Read Write As #nOut

    ' Filter each line
    On Error GoTo FilterTextError3
    Dim sLine As String, iLine As Long, eca As EChunkAction
    Dim fDoAll As Boolean
    Do Until EOF(nIn)
        Line Input #nIn, sLine
        iLine = iLine + 1
        If fDoAll = False Then
            eca = filter.Translate(sLine, iLine)
        End If
        Select Case eca
        Case ecaAbort
            Exit Do             ' Stop processing
        Case ecaTranslate
            Print #nOut, sLine  ' Write modified line to output
        Case ecaSkip
                                ' Ignore
        Case ecaTranslateAll
            Print #nOut, sLine  ' Write modified line to output
            fDoAll = True       ' Translate the rest without filter
        Case Else
            BugAssert True      ' Should never happen
        End Select
    Loop
    
    ' Close files
    On Error GoTo FilterTextError1
    Close nIn
    Close nOut
    If fReplace Then
        ' Destroy old file and replace it with new one
        Kill filter.Source
        On Error Resume Next   ' No more errors allowed
        Name sTarget As filter.Source
        ' If this fails, you're in trouble
        BugAssert Err = 0
    End If
    Exit Sub
        
FilterTextError3:
    Close nOut
FilterTextError2:
    Close nIn
FilterTextError1:
    ErrRaise Err
End Sub
'


' Applies filter to the IFilter.Source string and
' and saves the result in the IFilter.Target string.
Sub FilterText(filter As IFilter)
    Dim sSrc As String, sDst As String
    Dim iLine As Long, sLine As String
    
    sSrc = filter.Source
    sLine = MUtility.GetNextLine(sSrc)
    Do While sLine <> sEmpty
        ' Strip off sCrLf
        sLine = MUtility.RTrimLine(sLine)
        iLine = iLine + 1
        Select Case filter.Translate(sLine, iLine)
        Case ecaAbort:
            Exit Sub
        Case ecaTranslate:
            sDst = sDst & sLine & sCrLf
        Case ecaSkip:
        Case Else:
            BugAssert True
        End Select
        sLine = MUtility.GetNextLine
    Loop
    filter.Target = sDst
End Sub

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".Filter"
        Select Case e
        Case eeBaseFilter
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.EXEName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

