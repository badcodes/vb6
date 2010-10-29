VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GHtmlWeb"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const modHtmlWeb_SplitSymbol = "|"
Const modHtmlWeb_WebsiteDefaultFile = "·âÃæ|cover|Ê×Ò³|index|default|start|home|Ä¿Â¼|content|contents|aaa|bbb|00"

Public Function findDefaultHtml(ByRef sArrFilename() As String) As String

    Dim i As Integer
    Dim j As Integer
    Dim iArrUbound As Integer
    
    
    Dim sArrDefaultFilenameConst() As String
    Dim iArrDefaultFilenameConstUbound As Integer
    'Dim bGotIt As Boolean
    Dim sTmpFilename1 As String
    Dim stmpFilename2 As String

    If IsArray(sArrFilename) = False Then Exit Function
    sArrDefaultFilenameConst = Split(modHtmlWeb_WebsiteDefaultFile, modHtmlWeb_SplitSymbol)
    iArrUbound = UBound(sArrFilename)
    iArrDefaultFilenameConstUbound = UBound(sArrDefaultFilenameConst())

    For j = 0 To iArrDefaultFilenameConstUbound
        sTmpFilename1 = sArrDefaultFilenameConst(j)
        sTmpFilename1 = LCase$(sTmpFilename1)

        For i = 0 To iArrUbound
            stmpFilename2 = sArrFilename(i)
            stmpFilename2 = GetBaseName(stmpFilename2)
            stmpFilename2 = LCase$(stmpFilename2)

            If StrComp(sTmpFilename1, stmpFilename2, vbTextCompare) = 0 Then

                If slashCountInstr(sArrFilename(i)) < slashCountInstr(findDefaultHtml) _
                   Or findDefaultHtml = "" Then
                    findDefaultHtml = sArrFilename(i)

                    If slashCountInstr(findDefaultHtml) = 0 Then Exit Function
                End If

            End If

        Next

    Next

End Function

Public Function IsWebsiteDefaultFile(ByVal sFilename As String) As Boolean

    Dim i As Integer

    Dim sArrDefaultFilenameConst() As String
    Dim iArrDefaultFilenameConstUbound As Integer
    Dim sHtmlFilename As String
    sArrDefaultFilenameConst = Split(modHtmlWeb_WebsiteDefaultFile, modHtmlWeb_SplitSymbol)
    iArrDefaultFilenameConstUbound = UBound(sArrDefaultFilenameConst())

    For i = 0 To iArrDefaultFilenameConstUbound
        sHtmlFilename = sArrDefaultFilenameConst(i)
        sHtmlFilename = LCase$(sHtmlFilename)
        sFilename = GetBaseName(sFilename)
        sFilename = LCase$(sFilename)

        If sHtmlFilename = sFilename Then
            IsWebsiteDefaultFile = True
            Exit Function
        End If

    Next

End Function

Public Function getHtmlTitle(HtmlFile As String, testSize As Long) As String
Dim fn As Integer
Dim HeadText As String
fn = FreeFile()
'On Error GoTo toend
Open HtmlFile For Input As #fn
If testSize > LOF(fn) Or testSize < 1 Then testSize = LOF(fn)
HeadText = StrConv(InputB(testSize, #fn), vbUnicode)
Close (fn)
getHtmlTitle = linvblib.LeftRange(HeadText, "title>", "</", vbTextCompare, ReturnEmptyStr)
End Function

'Public Sub ParseHTML(HTML As String)
''We go through the HTML, character by character
''checking first for <, then for spaces, then
''quotation marks, and finally /. As we find
''them we fire events and continue parsing.
''
''Clean code with few relevant comments is better than
''unwieldy code commented to death, IMHO
''
'Dim IsValue, IsProperty, IsTag, RaisedTagBegin As Boolean
'Dim i As Long
'Dim C As String
'Dim CurrentProperty As String
'Dim CurrentPropertyValue As String
'Dim CurrentTag As String
'Dim CurrentText As String
''Remove tabs and returns, they have no place in HTML
'HTML = Replace(HTML, vbCrLf, "")
'HTML = Replace(HTML, vbTab, "")
''Start our searching
'For i = 1 To Len(HTML)
'    C = Mid(HTML, i, 1)
'    If IsTag = True Then
'        If IsProperty = True Then
'            If IsValue = True Then
'                If C = Chr(34) Then
'                    IsValue = False
'                    IsProperty = False
'                    CurrentPropertyValue = Trim(CurrentPropertyValue)
'                    CurrentProperty = Trim(CurrentProperty)
'                    RaiseEvent HTMLProperty(Left(CurrentProperty, Len(CurrentProperty) - 1), CurrentPropertyValue)
'                    CurrentPropertyValue = ""
'                    CurrentProperty = ""
'                Else
'                    CurrentPropertyValue = CurrentPropertyValue & C
'                End If
'            ElseIf C = Chr(34) Then
'                IsValue = True
'            Else
'                CurrentProperty = CurrentProperty & C
'            End If
'        Else
'            If C = " " Then
'                IsProperty = True
'                CurrentTag = Trim(CurrentTag)
'                CurrentTag = CurrentTag
'                If RaisedTagBegin = False Then
'                    RaiseEvent HTMLTagBegin(CurrentTag)
'                    RaisedTagBegin = True
'                End If
'            ElseIf C = ">" Then
'                IsTag = False
'                If Left(CurrentTag, 1) = "/" Then
'                    RaiseEvent HTMLTagClose(Right(CurrentTag, Len(CurrentTag) - 1))
'                ElseIf RaisedTagBegin = False Then
'                    RaiseEvent HTMLTagBegin(CurrentTag)
'                    RaiseEvent HTMLTagEnd(CurrentTag)
'                    RaisedTagBegin = True
'                Else
'                    RaiseEvent HTMLTagEnd(CurrentTag)
'                End If
'                CurrentTag = ""
'
'            Else
'                CurrentTag = CurrentTag & C
'            End If
'        End If
'    Else
'        If C = "<" Then
'            IsTag = True
'            RaisedTagBegin = False
'            If Trim(CurrentText) <> "" Then
'                RaiseEvent HTMLText(Trim(CurrentText))
'                CurrentText = ""
'            End If
'        Else
'            CurrentText = CurrentText & C
'        End If
'    End If
'Next i
'End Sub
Private Function skipChar(ByRef c As String, ByRef fNum As Integer) As String
Dim CC As String
Do While Not EOF(fNum)
    CC = Input$(1, #fNum)
    If CC <> c Then
        skipChar = CC
        Exit Function
    End If
Loop
End Function
Private Function skipUntil(ByRef c As String, ByRef fNum As Integer) As String
Dim CC As String
Do While Not EOF(fNum)
    CC = Input$(1, #fNum)
    If CC = c Then
        skipUntil = CC
        Exit Function
    End If
Loop
End Function

Public Function getTagsProperty(ByVal HtmlFile As String, ByVal tagName As String, ByVal propertyName As String, ByRef result() As String) As Long
'We go through the HTML, character by character
'checking first for <, then for spaces, then
'quotation marks, and finally /.

Dim c As String
Dim endChar As String
Dim Tag As String
Dim Property As String
Dim PropertyValue As String
Dim fNum As Integer

fNum = FreeFile()
Open HtmlFile For Input As fNum

'Remove tabs and returns, they have no place in HTML
Do Until EOF(fNum) = True
    Tag = ""
    'get Tag
    c = skipUntil("<", fNum)
    If c = "" Then Exit Do
    If EOF(fNum) Then Exit Do
    c = skipChar(" ", fNum)
    Do While EOF(fNum) = False
        Tag = Tag & c
        c = Input$(1, #fNum)
        If c = ">" Then GoTo LOOPLASTLINE
        If c = " " Then Exit Do
    Loop
    ' found Tag or tagName is empty
    ' get Property
    If LCase$(Tag) = LCase$(tagName) Or tagName = "" Then
        Debug.Print "<" & Tag;
        Do While c <> ">"
            Property = ""
            PropertyValue = ""
            endChar = ""
            c = ""
            c = skipChar(" ", fNum)
            If c = ">" Then Exit Do
            Do While Not EOF(fNum)
                Property = Property & c
                c = Input$(1, #fNum)
                If c = "=" Or c = ">" Then Exit Do
            Loop
            Debug.Print " " & Property;
            'get Property Value
            If c = "=" Then
                c = skipChar(" ", fNum)
                If c = Chr(34) Then
                    endChar = Chr(34)
                    c = ""
                Else
                    endChar = " "
                End If
                Do While Not EOF(fNum)
                    PropertyValue = PropertyValue & c
                    c = Input$(1, #fNum)
                    If c = endChar Or c = ">" Then Exit Do
                Loop
                Debug.Print "=" & PropertyValue;
             End If
            If PropertyValue <> "" And (LCase$(Property) = LCase$(propertyName) Or propertyName = "") Then
               ReDim Preserve result(0 To getTagsProperty)
               result(getTagsProperty) = PropertyValue
               getTagsProperty = getTagsProperty + 1
            End If
        Loop
        Debug.Print ">"
    End If
LOOPLASTLINE:
Loop
Close (fNum)
End Function



      

