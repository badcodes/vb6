Attribute VB_Name = "modMacros"
Option Explicit

'General
Public VBInstance              As VBIDE.VBE    ' The usual stuff by addins
Public Ini                     As New CIni     ' For saving settings
Public frmAddinDisplayed       As Boolean      ' frmAddIn stuff

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'constants for macro expantion:
Public Const CMB_CURSOR                                As String = "#CURSOR#"
Public Const CMB_STARTSEL                              As String = "#STARTSEL#"
Public Const CMB_ENDSEL                                As String = "#ENDSEL#"
Public Const CMB_LASTWORD                              As String = "#LASTWORD#"
Public Const CMB_DATE                                  As String = "#DATE#"
Public Const CMB_TIME                                  As String = "#TIME#"
Public Const CMB_PROCNAME                              As String = "#PROCNAME#"
Public Const CMB_PROCKIND                              As String = "#PROCKIND#"
Public Const CMB_PROCARG                               As String = "#PROCARG#"
Public Const CMB_PROCRETURNTYPE                        As String = "#PROCRETURNTYPE#"
Public Const CMB_PROCDESCRIPTION                       As String = "#PROCDESCRIPTION#"
Public Const CMB_MODULENAME                            As String = "#MODULENAME#"
Public Const CMB_MODULEFILENAME                        As String = "#MODULEFILENAME#"
Public Const CMB_MODULEFILEPATH                        As String = "#MODULEFILEPATH#"
Public Const CMB_MODULETYPE                            As String = "#MODULETYPE#"
Public Const CMB_PROJECTNAME                           As String = "#PROJECTNAME#"
Public Const CMB_PROJECTFILENAME                       As String = "#PROJECTFILENAME#"
Public Const CMB_PROJECTFILEPATH                       As String = "#PROJECTFILEPATH#"
Public Const CMB_PROJECTTYPE                           As String = "#PROJECTTYPE#"
Public Const CMB_INPUTBOX                              As String = "#INPUTBOX#"
Public Const CMB_PROGRAMMERNAME                        As String = "#PROGRAMMERNAME#"

'other macro stuff
Public Const MACRO_FILE                                As String = "\Macros.dat"

'codes and macros storage
Public mMacros() As String
Public mMacrosCount As Integer

'storage for procedure arguments.
Private mMacroArguments() As String

'storage for a macro parsed in single lines
Private mMacroLines() As String

'connects ExpandMacro with InsertMacro
Private bIsCursor As Boolean
Private bIsSelection As Boolean
Private bLastWord As Boolean

'store the words in a single line.
Private mMacroLineWords() As String

'preserve the the text in the macro text box if imported from
'the code window
Public IsFastMacro As Boolean

Public FormDisplayed As Boolean
Public hKinstalled As Boolean


Function lngProcedureType(oCM As CodeModule, startLine As Long) As Long
Dim ProcedureName As String
Dim strProcedureCode As String

lngProcedureType = -1

ProcedureName = oCM.ProcOfLine(startLine, vbext_pk_Proc Or vbext_pk_Get Or vbext_pk_Let Or vbext_pk_Set)

'
'   This all necessary due to the fact that we have to knwo the proper type of Get/Let to use
'   We are just getting one line now because we are doing this just to test and probe
On Error Resume Next

strProcedureCode = oCM.Lines(oCM.procStartLine(ProcedureName, vbext_pk_Proc), 1)

If Err.Number = 35 Then ' Sub or Function not defined
    Err.Clear
    strProcedureCode = oCM.Lines(oCM.procStartLine(ProcedureName, vbext_pk_Get), 1)

    If Err.Number = 35 Then ' Sub or Function not defined
        Err.Clear
        strProcedureCode = oCM.Lines(oCM.procStartLine(ProcedureName, vbext_pk_Let), 1)

        If Err.Number = 35 Then ' Sub or Function not defined
            Err.Clear
            strProcedureCode = oCM.Lines(oCM.procStartLine(ProcedureName, vbext_pk_Set), 1)

            If Err.Number = 0 Then
                lngProcedureType = vbext_pk_Set
            Else
' could not determine the type...
                lngProcedureType = -1
            End If
        Else
            lngProcedureType = vbext_pk_Let
        End If
    Else
        lngProcedureType = vbext_pk_Get
    End If

Else
    lngProcedureType = vbext_pk_Proc
End If

On Error GoTo 0
'
'   If we could not determine the type or it is not a type we wnat, skip on outta here
'If lngProcedureType <> -1 Then

End Function


Public Property Get ShiftPressed() As Boolean
ShiftPressed = (GetAsyncKeyState(vbKeyShift) <> 0)
End Property


Public Function ParseLines(ByVal instring As String) As Integer

Dim delimitList As String, oneChar As String, aWord As String, codeCount As Integer
Dim i As Integer, j As Integer, k As Integer

ReDim mMacroLines(0)
delimitList = Chr(13)

i = Len(instring)

For j = 1 To i
'Read one character at a time

    oneChar = VBA.Strings.Mid(instring, j, 1)
    k = InStr(delimitList, oneChar)
'Is this one a delimiter?
    If k = 0 Then
        aWord = aWord & oneChar
'If is isn't, add to the current word
    End If
    If k <> 0 Or j = i Then
'If it is, or if we're finished
        If aWord > vbNullString Then
            codeCount = codeCount + 1
            ReDim Preserve mMacroLines(codeCount)
            mMacroLines(codeCount - 1) = aWord
'Save new word
            aWord = vbNullString
        End If
    End If
Next j
ParseLines = UBound(mMacroLines)
'Return the array
End Function

Public Function CodeConvert(CodeToReplace As String) As String

Dim i As Integer, sTemp As String

For i = 0 To UBound(mMacros, 2) - 1
    'sTemp = mMacros(i).sCode
    sTemp = mMacros(1, i)
    If CodeToReplace = sTemp Then
        'CodeConvert = mMacros(i).sMacro
        CodeConvert = mMacros(2, i)
        Exit For
    End If
Next

End Function


Function FileExists(ByVal strPathName As String) As Boolean

'Returns: True if file exists, False otherwise

Dim hFile As Long
On Local Error Resume Next

'Remove any trailing directory separator character
If Right$(strPathName, 1) = "\" Then
    strPathName = Left$(strPathName, Len(strPathName) - 1)
End If

'Attempt to open the file, return value of
'this function is False if an error occurs
'on open, True otherwise
hFile = FreeFile
Open strPathName For Input As hFile

FileExists = Err = 0

Close hFile
Err = 0

End Function

Public Sub InsertMacro()

Dim oCM As CodeModule
Dim oCP As CodePane
Dim sMCode As String
Dim sLine As String
Dim lsLine As Long, leLine As Long
Dim lsCol As Long, leCol As Long

Dim L As Long, M As Long, n As Long, PosBefore As Long, PosAfter As Long
Dim sMacro As String
Dim sNewLine As String, sLineHead As String, sLineFoot As String, sLastWord As String
Dim sTemp As String
Dim lCursorColumn As Long
Dim lCursorLine As Long
Dim SelectionFailed As Boolean
Dim lSelLineStart As Long, lSelColStart As Long, lSelLineEnd As Long, lSelColEnd As Long


Const cReplaceTab = "    "

'Reference the codemodule and codepane
Set oCM = VBInstance.ActiveCodePane.CodeModule
Set oCP = oCM.CodePane

With oCP
    .GetSelection lsLine, lsCol, leLine, leCol 'Find out what is selected
    sLine = oCM.Lines(lsLine, 1) 'Get the line were on
End With

If lsCol <= 2 Then Exit Sub '1 space can't be a macro

PosBefore = InStrRev(sLine, " ", leCol - 2) 'get the word start, skipping space
If PosBefore Then
    sLineHead = Mid(sLine, 1, PosBefore) 'including space
    sLine = Replace$(sLine, sLineHead, vbNullString, 1, 1) 'including space
End If

PosAfter = InStr(1, sLine, " ")
If PosAfter Then
    sLineFoot = Mid(sLine, PosAfter + 1, Len(sLine) - PosAfter) 'no space
    sLine = Replace$(sLine, sLineFoot, vbNullString) 'including space
End If

sMacro = CodeConvert(Left(sLine, Len(sLine) - 1)) 'cut the last chr, the space

If Len(sMacro) Then 'if the macro code is a valid one
    sMacro = ExpandMacro(sMacro, lsLine) 'expand also the embedded macros

    If bLastWord Then
        M = InStrRev(sLineHead, " ", Len(sLineHead) - 1) 'skip the space

        If M = 0 Then 'no space, just one word
            sLastWord = Left(sLineHead, Len(sLineHead) - 1)
        Else 'pick up the last word
            sLastWord = Mid$(sLineHead, M + 1, Len(sLineHead) - M - 1)
        End If

        sMacro = Replace$(sMacro, CMB_LASTWORD, sLastWord)
    End If

    sNewLine = sLineHead & sMacro & sLineFoot

    If bIsCursor Then
        sTemp = sNewLine
        sTemp = Replace$(sTemp, vbKeyTab, cReplaceTab) 'the IDE doesn't like tabs
        L = InStr(1, sTemp, Chr(27)) 'where is the cursor mark?
        M = InStr(1, sTemp, Chr(13)) 'the end of the first line

        If (L > M) And (M > 0) Then 'Chr(27) is not in the first line
            sTemp = Left(sTemp, L) 'cut the string, the rest is irrelevant

            Do 'because we are just looking for cursor's line and column:
                M = InStr(1, sTemp, Chr(13))
                If M = 0 Then Exit Do
                lCursorLine = lCursorLine + 1
                sTemp = Right(sTemp, Len(sTemp) - M) 'it will be the line containing chr(27)
            Loop

            lCursorLine = lCursorLine + lsLine
            lCursorColumn = InStr(1, sTemp, Chr(27))
            oCM.ReplaceLine lsLine, sNewLine
            oCP.SetSelection lCursorLine, lCursorColumn - 1, lCursorLine, lCursorColumn
            SendKeys Chr(8)
        Else  'Chr(27) is in the first line
            lCursorLine = lsLine
            lCursorColumn = InStr(1, sNewLine, Chr(27))
            oCM.ReplaceLine lsLine, sNewLine
            oCP.SetSelection lCursorLine, lCursorColumn, lCursorLine, lCursorColumn + 1
            SendKeys Chr(8)
        End If
    ElseIf bIsSelection Then 'an embedded selection
        sTemp = sNewLine
        sTemp = Replace$(sTemp, vbKeyTab, cReplaceTab) 'the IDE doesn't like tabs
        L = InStr(1, sTemp, Chr(28)) 'where is the endsel mark?
        M = InStr(1, sTemp, Chr(13)) 'the end of the first line

        If (L > M) And (M > 0) Then 'Chr(28) is not in the first line
            sTemp = Left(sTemp, L) 'cut the string, the rest is irrelevant

            Do 'because we are just looking for cursor's line and column:
                M = InStr(1, sTemp, Chr(13))
                If M = 0 Then Exit Do
                lSelLineEnd = lSelLineEnd + 1
                sTemp = Right(sTemp, Len(sTemp) - M) 'it will be the line containing chr(28)
            Loop

            n = InStr(1, sTemp, Chr(27)) 'get start col

'we don't support multiline selection
            If n = 0 Then 'just print with no selection
                sNewLine = Replace$(sNewLine, Chr(27), vbNullString)
                sNewLine = Replace$(sNewLine, Chr(28), vbNullString)
                oCM.ReplaceLine lsLine, sNewLine
            Else 'single line selection, we endorse it
                lSelLineEnd = lSelLineEnd + lsLine
                lSelLineStart = lSelLineEnd

                lSelColStart = n - 1
                lSelColEnd = InStr(1, sTemp, Chr(28)) - 1 'we have to cut out Chr(27) and Chr(28)

                sNewLine = Replace$(sNewLine, Chr(27), vbNullString)
                sNewLine = Replace$(sNewLine, Chr(28), vbNullString)
                oCM.ReplaceLine lsLine, sNewLine
                oCP.SetSelection lSelLineStart, lSelColStart, lSelLineEnd, lSelColEnd
            End If
        Else  'Chr(28) is in the first line
            lSelLineEnd = lsLine
            lSelLineStart = lSelLineEnd

            lSelColStart = InStr(1, sTemp, Chr(27))
            lSelColEnd = InStr(1, sTemp, Chr(28)) - 1 'we have to cut out Chr(27) and Chr(28)

            sNewLine = Replace$(sNewLine, Chr(27), vbNullString)
            sNewLine = Replace$(sNewLine, Chr(28), vbNullString)
            oCM.ReplaceLine lsLine, sNewLine
            oCP.SetSelection lSelLineStart, lSelColStart, lSelLineEnd, lSelColEnd
        End If
    Else 'no cursor, no selection
        oCM.ReplaceLine lsLine, sNewLine 'life is easy
        SendKeys "{Up}{End}"
    End If
End If

End Sub
Public Function GetArguments(oCM As CodeModule, startLine As Long) As Integer
' if GetArguments returns null the array is loaded.
On Error GoTo eH

Dim ProcString As String, ProcedureStartLine As Long, ProcedureBodyLine As Long
Dim ProcStringLines As Long, ProcedureName As String, ProcedureKind As Long
Const cBrackLeft = "("
Const cBrackRight = ")"
Const cOpt = "Optional "
Const cVal = "ByVal "
Const cRef = "ByRef "
Const cAs = " As "
Const cAs1 = " As"
Const cComma = ","
Dim sTemp As String, strWord As String
Dim i As Integer, j As Integer, k As Integer
Dim countArguments As Integer

ProcedureName = oCM.ProcOfLine(startLine, vbext_pk_Proc Or vbext_pk_Get Or vbext_pk_Let Or vbext_pk_Set)
ProcedureKind = lngProcedureType(oCM, startLine)
ProcedureStartLine = oCM.procStartLine(ProcedureName, ProcedureKind) + 1

ProcString = oCM.Lines(ProcedureStartLine, 1)
i = InStr(1, ProcString, " _")

Do While i > 0
    ProcString = Replace$(ProcString, " _", " ")
    j = j + 1
    ProcString = ProcString & oCM.Lines(ProcedureStartLine + j, 1)
    i = InStr(ProcString, " _")
Loop

i = InStr(ProcString, cBrackLeft)
j = InStr(ProcString, cBrackRight)

If j = i + 1 Then
    GetArguments = 0
    ReDim mMacroArguments(0)
    Exit Function
End If

sTemp = Mid(ProcString, i + 1, j - i - 1)
k = InStr(sTemp, cOpt)

Do While k > 0
    sTemp = Left(sTemp, k - 1) & Right(sTemp, Len(sTemp) - k - 8)
    k = InStr(sTemp, cOpt)
Loop

k = InStr(sTemp, cVal)

Do While k > 0
    sTemp = Left(sTemp, k - 1) & Right(sTemp, Len(sTemp) - k - 5)
    k = InStr(sTemp, cVal)
Loop

k = InStr(sTemp, cRef)

Do While k > 0
    sTemp = Left(sTemp, k - 1) & Right(sTemp, Len(sTemp) - k - 5)
    k = InStr(sTemp, cRef)
Loop

'find
i = InStr(sTemp, cAs)
strWord = Mid(sTemp, 1, i - 1)
'clean
strWord = Replace$(strWord, Chr(13), vbNullString)
strWord = Replace$(strWord, Chr(10), vbNullString)
strWord = Replace$(strWord, "_", vbNullString)
'store
countArguments = countArguments + 1
ReDim mMacroArguments(countArguments)
mMacroArguments(countArguments - 1) = strWord

'cut out
j = InStr(sTemp, cComma)
sTemp = Right(sTemp, Len(sTemp) - j + 1)

Do While j > 0
'find
    i = InStr(sTemp, cComma)
    j = InStr(sTemp, cAs1)
    strWord = Mid(sTemp, i + 2, j - i - 2)
'clean
    strWord = Replace$(strWord, Chr(13), vbNullString)
    strWord = Replace$(strWord, Chr(10), vbNullString)
'store
    countArguments = countArguments + 1
    ReDim Preserve mMacroArguments(countArguments)
    mMacroArguments(countArguments - 1) = strWord

'cut out
    j = InStr(2, sTemp, cComma)
    sTemp = Right(sTemp, Len(sTemp) - j + 1)
Loop

GetArguments = UBound(mMacroArguments)
Exit Function
eH:
GetArguments = 0
End Function


Public Function ExpandMacro(sMacro As String, startLine As Long) As String
Const cNoArg = "No Argument" & vbCrLf

Dim i As Integer, StartPos As Integer, EndPos As Integer, iTemp As Integer
Dim sCopyOfMacro As String, sTemp As String, sPathTmp As String, sNameTmp As String
Dim Str1 As String, Str2 As String, Str3 As String, Str4 As String
Dim lsLine As Long, lsCol As Long, leLine As Long, leCol As Long
Dim oCM As CodeModule
Dim oCP As CodePane
Dim oProj As VBProject
Dim oComp As VBComponent

'Reference
Set oCM = VBInstance.ActiveCodePane.CodeModule
Set oCP = oCM.CodePane
Set oProj = VBInstance.ActiveVBProject
Set oComp = VBInstance.ActiveCodePane.VBE.SelectedVBComponent

bIsCursor = False
bIsSelection = False
bLastWord = False

sCopyOfMacro = sMacro

StartPos = InStr(1, sCopyOfMacro, CMB_LASTWORD)
If StartPos > 0 Then
    bLastWord = True
End If

StartPos = InStr(1, sCopyOfMacro, CMB_DATE)
If StartPos > 0 Then
    sTemp = Date
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_DATE, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_TIME)
If StartPos > 0 Then
    sTemp = Time
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_TIME, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_TIME)
If StartPos > 0 Then
    sTemp = Time
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_TIME, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_PROCNAME)
If StartPos > 0 Then
    sTemp = oCM.ProcOfLine(startLine, vbext_pk_Proc Or vbext_pk_Get Or vbext_pk_Let Or vbext_pk_Set)
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_PROCNAME, sTemp)
End If 'ok

StartPos = InStr(1, sCopyOfMacro, CMB_PROJECTNAME)
If StartPos > 0 Then
    sTemp = oProj.Name
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_PROJECTNAME, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_PROJECTFILENAME)
If StartPos > 0 Then
    sTemp = oProj.FileName
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_PROJECTFILENAME, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_PROJECTFILEPATH)
If StartPos > 0 Then
    sTemp = App.Path
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_PROJECTFILEPATH, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_PROGRAMMERNAME)
If StartPos > 0 Then
    sTemp = frmAddIn.WLText3.Text
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_PROGRAMMERNAME, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_PROJECTTYPE)
If StartPos > 0 Then
    sTemp = oProj.Type
    Select Case oProj.Type
    Case 0
        sTemp = "Standard Executable"
    Case 1
        sTemp = "ActiveX Executable"
    Case 2
        sTemp = "ActiveX DLL"
    Case 3
        sTemp = "ActiveX OCX"
    End Select
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_PROJECTTYPE, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_MODULEFILENAME)
If StartPos > 0 Then
    sTemp = oComp.FileNames(1)
    SplitPathname sTemp, sPathTmp, sNameTmp
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_MODULEFILENAME, sNameTmp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_MODULEFILEPATH)
If StartPos > 0 Then
    sTemp = oComp.FileNames(1)
    SplitPathname sTemp, sPathTmp, sNameTmp
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_MODULEFILEPATH, sPathTmp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_MODULENAME)
If StartPos > 0 Then
    sTemp = oComp.Name
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_MODULENAME, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_MODULETYPE)
If StartPos > 0 Then
    iTemp = oComp.Type

    Select Case iTemp
    Case 11
        sTemp = "ActiveXDesigner"
    Case 2
        sTemp = "ClassModule"
    Case 9
        sTemp = "DocObject"
    Case 3
        sTemp = "MSForm"
    Case 7
        sTemp = "PropPage"
    Case 10
        sTemp = "RelatedDocument"
    Case 4
        sTemp = "ResFile"
    Case 1
        sTemp = "StdModule"
    Case 8
        sTemp = "UserControl"
    Case 5
        sTemp = "VBForm"
    Case 6
        sTemp = "VBMDIForm"
    End Select

    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_MODULETYPE, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_PROCDESCRIPTION)
If StartPos > 0 Then
    sTemp = oComp.Description
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_PROCDESCRIPTION, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_PROCKIND)
If StartPos > 0 Then

    sTemp = oCM.ProcOfLine(startLine, vbext_pk_Proc Or vbext_pk_Get Or vbext_pk_Let Or vbext_pk_Set)

    If InStr(1, sTemp, "Function ") > 0 Then
        sTemp = "Function"
    ElseIf InStr(1, sTemp, "Sub ") > 0 Then
        sTemp = "Sub"
    ElseIf InStr(1, sTemp, "Property Get") > 0 Then
        sTemp = "Property Get"
    ElseIf InStr(1, sTemp, "Property Let") > 0 Then
        sTemp = "Property Let"
    End If
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_PROCKIND, vbNullString)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_PROCRETURNTYPE)
If StartPos > 0 Then
    Str1 = oCM.ProcOfLine(startLine, vbext_pk_Proc Or vbext_pk_Get Or vbext_pk_Let Or vbext_pk_Set)
    i = InStr(1, Str1, ")")

    If Len(Str1) > i Then
        i = InStr(i, Str1, "As ")
        sTemp = Right(Str1, Len(Str1) - i - 2)
    End If

    i = Len(sTemp)
    If i = 0 Then sTemp = "NO RETURN"
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_PROCRETURNTYPE, sTemp)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_PROCARG)
If StartPos > 0 Then
'store from the line with  CMB_PROCARG what is before CMB_PROCARG,
'it must be printed before each variable
    Str1 = Left(sCopyOfMacro, StartPos - 1)
    i = InStrRev(Str1, Chr(13))

    If i > 0 Then
        Str1 = Right(Str1, Len(Str1) - i - 1)

'store the head of the macro, until the line with CMB_PROCARG.
'Such line is excluded.
        sTemp = Left(sCopyOfMacro, i - 1) & vbCrLf
    End If

    i = InStr(StartPos + 9, sCopyOfMacro, Chr(13))

    If i > 0 Then
        Str3 = Mid(sCopyOfMacro, StartPos + 9, i - (StartPos + 9))
    End If

    i = InStr(StartPos + 9, sCopyOfMacro, Chr(13))
    If i > 0 Then Str4 = Mid(sCopyOfMacro, i + 1)


    If GetArguments(oCM, startLine) = 0 Then
        Str2 = cNoArg
        sTemp = sTemp & Str1 & Str2 & Str4
        sCopyOfMacro = sTemp
    Else
        For i = 0 To UBound(mMacroArguments) - 1
            sTemp = sTemp & Str1 & mMacroArguments(i) & Str3 & vbCrLf
        Next


        sTemp = Left(sTemp, Len(sTemp) - 1) & Str4
        sCopyOfMacro = sTemp
    End If
'Debug.Print sCopyOfMacro
End If

''''''''''''''''''

StartPos = InStr(1, sCopyOfMacro, CMB_INPUTBOX)
If StartPos > 0 Then
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_INPUTBOX, vbNullString)
End If

StartPos = InStr(1, sCopyOfMacro, CMB_CURSOR)
If StartPos > 0 Then
    bIsCursor = True
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_CURSOR, Chr(27))
End If

StartPos = InStr(1, sCopyOfMacro, CMB_STARTSEL)
If StartPos > 0 Then
    sCopyOfMacro = Replace$(sCopyOfMacro, Chr(27), vbNullString) 'user mistakes...
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_STARTSEL, Chr(27))
    sCopyOfMacro = Replace$(sCopyOfMacro, CMB_ENDSEL, Chr(28))
    bIsSelection = True
    bIsCursor = False
End If

ExpandMacro = sCopyOfMacro

End Function
Public Sub SplitPathname(ByVal fullname$, fpath$, FName$)
Dim i%, p%
On Error Resume Next
Do
    p = i
    i = InStr(i + 1, fullname, "\")
Loop While i
If p Then
    fpath = Left$(fullname, p)
End If
FName = Right$(fullname, Len(fullname) - p)
End Sub




