Attribute VB_Name = "MgCHtmlWeb"
Option Explicit
Private Const hw_SplitFlag = "|"
Private Const hw_DefaultFilename = "·âÃæ|cover|Ê×Ò³|index|default|start|home|Ä¿Â¼|content|contents|aaa|bbb|00"

Public Function findDefaultHtml(ByRef arrFileName() As String, Optional sSearchName As String = "") As String

    Dim i As Integer
    Dim j As Integer
    Dim iEnd As Integer
    Dim jEnd As Integer
    Dim sTmp1 As String
    Dim sTmp2 As String

    Dim arrSearchName() As String
    Dim fso As New gCFileSystem
    Dim hString As New gCString
    
    If IsArray(arrFileName) = False Then Exit Function
    
    If sSearchName <> "" Then sSearchName = sSearchName & hw_SplitFlag
    sSearchName = sSearchName & hw_DefaultFilename
    
    arrSearchName = Split(sSearchName, hw_SplitFlag)
    
    iEnd = UBound(arrSearchName)
    jEnd = UBound(arrFileName)

    For i = 0 To iEnd
        sTmp1 = arrSearchName(i)
        For j = 0 To jEnd
            sTmp2 = arrFileName(j)
            sTmp2 = fso.GetBaseName(sTmp2)
            If StrComp(sTmp1, sTmp2, vbTextCompare) = 0 Then
            
                If findDefaultHtml = "" Or _
                    hString.slashCountInstr(arrFileName(j)) < _
                    hString.slashCountInstr(findDefaultHtml) _
                    Then
                    findDefaultHtml = arrFileName(j)
                    If hString.slashCountInstr(findDefaultHtml) = 0 Then Exit Function
                End If
                
            End If
        Next
    Next

End Function

Public Function IsWebsiteDefaultFile(ByVal sFileName As String) As Boolean

    Dim i As Integer
    Dim fso As New gCFileSystem
    Dim sArrDefaultFilenameConst() As String
    Dim iArrDefaultFilenameConstUbound As Integer
    Dim sHtmlFilename As String
    sArrDefaultFilenameConst = Split(hw_DefaultFilename, hw_SplitFlag)
    iArrDefaultFilenameConstUbound = UBound(sArrDefaultFilenameConst())

    For i = 0 To iArrDefaultFilenameConstUbound
        sHtmlFilename = sArrDefaultFilenameConst(i)
        sHtmlFilename = LCase$(sHtmlFilename)
        sFileName = fso.GetBaseName(sFileName)
        sFileName = LCase$(sFileName)

        If sHtmlFilename = sFileName Then
            IsWebsiteDefaultFile = True
            Exit Function
        End If
    Next
End Function

Public Function getHtmlTitle(HtmlFile As String) As String
'<EhHeader>
On Error GoTo getHtmlTitle_EXIT
'</EhHeader>

Dim c As String
Dim fNum As Integer
Dim sTag As String


fNum = FreeFile()
'On Error GoTo toend
Open HtmlFile For Input As #fNum
Do While Not EOF(fNum)
    If SkipUntil("<", fNum) Then c = Read(fNum)
    'skipChar " ", fNum
    sTag = StrUntil(" >", fNum)
    sTag = LCase$(sTag)
    If sTag = "title" Then
        If SkipUntil(">", fNum) Then c = Read(fNum)
        getHtmlTitle = StrUntil("<", fNum)
        If EOF(fNum) Then getHtmlTitle = ""
        Exit Function
    End If
Loop
Close (fNum)
'<EhFooter>
Exit Function
getHtmlTitle_EXIT:
'</EhFooter>
End Function

'Public Function getTagsProperty(ByVal HtmlFile As String, ByVal tagName As String, ByVal propertyName As String, ByRef result() As String) As Long
''We go through the HTML, character by character
''checking first for <, then for spaces, then
''quotation marks, and finally /.
'Dim lcFile As New gCFile
'Dim C As String
'Dim endChar As String
'Dim Tag As String
'Dim Property As String
'Dim PropertyValue As String
'Dim fNum As Integer
'
'fNum = FreeFile()
'Open HtmlFile For Input As fNum
'
''Remove tabs and returns, they have no place in HTML
'Do Until EOF(fNum) = True
'    Tag = ""
'    'get Tag
'    C = lcFile.skipUntil("<", fNum)
'    If C = "" Then Exit Do
'    If EOF(fNum) Then Exit Do
'    C = lcFile.skipChar(" ", fNum)
'    Do While EOF(fNum) = False
'        Tag = Tag & C
'        C = Input$(1, #fNum)
'        If C = ">" Then GoTo LOOPLASTLINE
'        If C = " " Then Exit Do
'    Loop
'    ' found Tag or tagName is empty
'    ' get Property
'    If LCase$(Tag) = LCase$(tagName) Or tagName = "" Then
'        Debug.Print "<" & Tag;
'        Do While C <> ">"
'            Property = ""
'            PropertyValue = ""
'            endChar = ""
'            C = ""
'            C = lcFile.skipChar(" ", fNum)
'            If C = ">" Then Exit Do
'            Do While Not EOF(fNum)
'                Property = Property & C
'                C = Input$(1, #fNum)
'                If C = "=" Or C = ">" Then Exit Do
'            Loop
'            Debug.Print " " & Property;
'            'get Property Value
'            If C = "=" Then
'                C = lcFile.skipChar(" ", fNum)
'                If C = Chr(34) Then
'                    endChar = Chr(34)
'                    C = ""
'                Else
'                    endChar = " "
'                End If
'                Do While Not EOF(fNum)
'                    PropertyValue = PropertyValue & C
'                    C = Input$(1, #fNum)
'                    If C = endChar Or C = ">" Then Exit Do
'                Loop
'                Debug.Print "=" & PropertyValue;
'             End If
'            If PropertyValue <> "" And (LCase$(Property) = LCase$(propertyName) Or propertyName = "") Then
'               ReDim Preserve result(0 To getTagsProperty)
'               result(getTagsProperty) = PropertyValue
'               getTagsProperty = getTagsProperty + 1
'            End If
'        Loop
'        Debug.Print ">"
'    End If
'LOOPLASTLINE:
'Loop
'Close (fNum)
'End Function
