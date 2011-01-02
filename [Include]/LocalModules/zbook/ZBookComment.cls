VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZhComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public sTitle As String
Public sAuthor As String
Public sPublisher As String
Public sDate As String
Public sCatalog As String
Public sDefaultfile As String
Public sHHCFile As String
Public sContentFile As String
Public zvShowLeft As zhtmVisiablity
Public zvShowMenu As zhtmVisiablity
Public zvShowStatusBar As zhtmVisiablity
Public ContentName As New CStringArray ' CStringCollection
Public ContentLocation As New CStringArray
Private mLContentcount As Long
'Private sArrcontent() As String '´Ó0¿ªÊ¼
Private bZhLoaded As Boolean
Const ContentSeparator1 = "|"
Const ContentSeparator2 = ","

Public Sub EraseObjects()

    Set ContentName = Nothing
    Set ContentLocation = Nothing

End Sub

Public Property Get lContentCount() As Long

    lContentCount = mLContentcount

End Property

Public Property Get ToString() As String

    ToString = InfoText & vbCrLf & ContentText

End Property

Public Property Get InfoText() As String

    Dim sLong As New CAppendString
    sLong.Append "[Info]" & vbCrLf
    sLong.Append "Title=" & sTitle & vbCrLf
    sLong.Append "Author=" & sAuthor & vbCrLf
    sLong.Append "Date=" & sDate & vbCrLf
    sLong.Append "Publisher=" & sPublisher & vbCrLf
    sLong.Append "Catalog=" & sCatalog & vbCrLf
    sLong.Append "DefaultFile=" & sDefaultfile & vbCrLf
    sLong.Append "HHCfile=" & sHHCFile & vbCrLf
    sLong.Append "ContentFile=" & sContentFile & vbCrLf
    sLong.Append "[Style]" & vbCrLf
    sLong.Append "ShowLeft=" & zvShowLeft & vbCrLf
    sLong.Append "ShowMenu=" & zvShowMenu & vbCrLf
    sLong.Append "ShowStatusBar=" & zvShowStatusBar
    InfoText = sLong.Value
    Set sLong = Nothing

End Property

Public Property Get ContentText() As String

    Dim sLong As New CAppendString
    '    Dim i As Long, m As Long

    If mLContentcount < 1 Then Exit Property
    '    m = mLContentcount - 1
    '
    sLong.Append "[Content]"
   Dim i As Long
   Dim iEnd As Long
   iEnd = ContentName.Count
   
   For i = 1 To iEnd
    sLong.Append vbCrLf & ContentName.Item(i) & ContentSeparator1 & ContentLocation.Item(i)
   Next
    
    '
    '    For i = 1 To m
    '        sLong.AppendLine ContentName(i) & ContentSeparator1 & ContentLocation(i)
    '    Next
    '
    '    sLong.Append ContentName(mLContentcount) & ContentSeparator1 & ContentLocation(mLContentcount)
    ContentText = sLong.Value
    Set sLong = Nothing

End Property

'
'Public Function readContentInIndex(lContentIndex As Long, sContentName As String, sContentLocation As String) As Boolean
'
'    readContentInIndex = False
'
'    If bZhLoaded = False Then Exit Function
'
'    If mLContentcount = 0 Then Exit Function
'
'    If lContentIndex > mLContentcount - 1 Then Exit Function
'    sContentName = sArrcontent(0, lContentIndex)
'    sContentLocation = sArrcontent(1, lContentIndex)
'    readContentInIndex = True
'
'End Function
Public Sub CopyContentTo(sArrDestContent() As String)

    If mLContentcount <= 0 Then Exit Sub
    Dim i As Long
    ReDim sArrDestContent(1, mLContentcount - 1) As String

    For i = 1 To mLContentcount
        sArrDestContent(0, i - 1) = ContentName.Item(i)
        sArrDestContent(1, i - 1) = ContentLocation.Item(i)
    Next

End Sub

'Public Sub CopyContentFrom(sArrSourceContent() As String)
'
'    mLContentcount = UBound(sArrSourceContent, 2)
'    ReDim sArrcontent(1, mLContentcount) As String
'    sArrcontent() = sArrSourceContent()
'    mLContentcount = mLContentcount + 1
'
'End Sub
Public Sub selfReset()

    sTitle = ""
    sAuthor = ""
    sPublisher = ""
    sDefaultfile = ""
    sDate = ""
    sCatalog = ""
    sHHCFile = ""
    sContentFile = ""
    zvShowLeft = zhtmVisiableDefault
    zvShowMenu = zhtmVisiableDefault
    zvShowStatusBar = zhtmVisiableDefault
    mLContentcount = 0
    Set ContentName = New CStringArray
    Set ContentLocation = New CStringArray
    'Erase sArrcontent()
    bZhLoaded = False

End Sub

Public Function parseHHC(ByRef sPathname As String, Optional ByVal sBasePath As String = "") As Boolean

    Dim HHCText As String
    Dim fNum As Integer
    fNum = FreeFile
    Open sPathname For Binary Access Read As #fNum
    HHCText = String$(LOF(fNum), " ")
    Get fNum, , HHCText
    Close #fNum
    parseHHC = parseHHCText(HHCText, sBasePath)

End Function

Public Function parseHHCText(ByRef HHCText As String, Optional ByVal sBasePath As String = "") As Boolean

    Dim hdoc As New HTMLDocument
'FIXIT: Declare 'ThisChild' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
    Dim ThisChild As Object
    Dim sAll() As String
    Dim i As Long
    Dim aCount As Long
    hdoc.body.innerHTML = HHCText
    mLContentcount = 0
    Set ContentName = New CStringArray
    Set ContentLocation = New CStringArray

    For Each ThisChild In hdoc.body.childNodes

        If ThisChild.nodeName = "UL" Then getLI ThisChild, sAll, aCount, 0, "", sBasePath
    Next

    'mLContentcount = mLContentcount + 1

    For i = 1 To aCount
        If sAll(0, i) <> "" Then
            mLContentcount = mLContentcount + 1
            ContentName.Add sAll(0, i)
            ContentLocation.Add linvblib.UnescapeUrl(sAll(1, i))
        End If
    Next

    Set hdoc = Nothing
    Set ThisChild = Nothing

End Function

Public Function parseZhComment(ByRef strCMTFilePath As String) As Boolean

    Dim fNum As Integer
    Dim sText As String
    fNum = FreeFile
    Open strCMTFilePath For Binary Access Read As #fNum
    sText = String$(LOF(fNum), " ")
    Get #fNum, , sText
    Close fNum
    parseZhComment = parseZhCommentText(sText)

End Function

'Public Function saveZhCommentToFile(strCMTFilePath As String) As Boolean
'
'    Dim fso As New scripting.FileSystemObject
'    Dim fsoTS As scripting.TextStream
'    Dim i As Long
'    Set fsoTS = fso.CreateTextFile(strCMTFilePath, True)
'
'    With fsoTS
'        .WriteLine "[Info]"
'        .WriteLine "Title=" & sTitle
'        .WriteLine "Author=" & sAuthor
'        .WriteLine "Date=" & sDate
'        .WriteLine "Publisher=" & sPublisher
'        .WriteLine "Catalog=" & sCatalog
'        .WriteLine "DefaultFile=" & sDefaultfile
'        .WriteLine "[Style]"
'        .WriteLine "ShowLeft=" & zvShowLeft
'        .WriteLine "ShowMenu=" & zvShowMenu
'        .WriteLine "ShowStatusBar=" & zvShowStatusBar
'        .WriteLine "[Content]"
'        Dim lEnd As Long
'        lEnd = mLContentcount - 1
'
'        For i = 0 To lEnd
'            .WriteLine contentname(i) & ContentSeparator1 & contentLocation(i)
'        Next
'
'    End With
'
'    fsoTS.Close
'
'End Function
Public Function parseZhCommentText(ByRef sCMT As String) As Boolean

    Dim cmtLine As String
    Dim pos As Integer
    Dim bContentStart As Boolean
    Dim sIniName As String
    Dim sIniValue As String
    Dim sCMTLine() As String
    Dim lCur As Long
    Dim lLastLine As Long
    Dim cName As String
    Dim cLocation As String
    parseZhCommentText = False
    bContentStart = False
    mLContentcount = 0
    Set ContentName = New CStringArray
    Set ContentLocation = New CStringArray
    sCMT = Replace$(sCMT, vbCrLf, vbLf)
    sCMTLine = Split(sCMT, vbLf)
    lLastLine = UBound(sCMTLine)
    
    
    For lCur = 0 To lLastLine
        cmtLine = sCMTLine(lCur)
        'Check if content start by compare cmtline with "[CONTENT]"

        If (Not bContentStart) And StrComp(LCase$(Trim$(cmtLine)), "[content]", vbTextCompare) = 0 Then
            bContentStart = True
            GoTo forContinue
        End If

        If bContentStart And cmtLine <> "" Then
            'ReDim Preserve sArrcontent(1, mLContentcount) As String
            pos = InStrRev(cmtLine, ContentSeparator1)

            If pos <= 0 Then pos = InStrRev(cmtLine, ContentSeparator2)

            If pos > 1 Then
                cName = Left$(cmtLine, pos - 1)
                cLocation = Right$(cmtLine, Len(cmtLine) - pos)
                If cName <> "" Then
                ContentName.Add Left$(cmtLine, pos - 1)
                ContentLocation.Add Right$(cmtLine, Len(cmtLine) - pos)
                mLContentcount = mLContentcount + 1
                End If
'            Else
'                ContentName.Add cmtLine
'                ContentLocation.Add ""
            End If

            
        ElseIf InStr(cmtLine, "=") > 0 Then
            pos = InStr(cmtLine, "=")
            sIniValue = Right$(cmtLine, Len(cmtLine) - pos)
            sIniName = LCase$(Left$(cmtLine, pos - 1))

            Select Case LCase$(sIniName)
            Case "showleft"
                zvShowLeft = Val(sIniValue)
            Case "showmenu"
                zvShowMenu = Val(sIniValue)
            Case "showstatusbar"
                zvShowStatusBar = Val(sIniValue)
            Case "defaultfile"
                sDefaultfile = sIniValue
            Case "author"
                sAuthor = sIniValue
            Case "title"
                sTitle = sIniValue
            Case "publisher"
                sPublisher = sIniValue
            Case "hhcfile"
                sHHCFile = sIniValue
            Case "contentfile"
                sContentFile = sIniValue
            Case "date"
                sDate = sIniValue
            Case "catalog"
                sCatalog = sIniValue
            End Select

        End If

forContinue:
    Next

    bZhLoaded = True

End Function

'Private Function HhcToZhc(sPathname As String) As String
'
'h = "C:\1.HHC"
'
'Dim HHCText As String
'Dim hdoc As New HTMLDocument
'Dim ThisChild As Object
'Dim sAll() As String
'Dim LCount As Integer
'Dim fNUM As Integer
'
'fNUM = FreeFile
'Open sPathname For Binary Access Read As #fNUM
'HHCText = String(LOF(fNUM), " ")
'Get fNUM, , HHCText
'Close #fNUM
'
'hdoc.body.innerHTML = HHCText
'
'ReDim sAll(1, 0) As String
'
'For Each ThisChild In hdoc.body.childNodes
'    If ThisChild.nodeName = "UL" Then getLI ThisChild, sAll, LCount, ""
'Next
'
'
''For i = 1 To LCount
''Debug.Print sAll(0, i) & "=" & sAll(1, i)
''Next
'
'End Function
Private Sub getLI(ByVal ULE As HTMLUListElement, ByRef sAll() As String, ByRef iStart As Long, ByVal iParent As Long, ByVal sParent As String, Optional ByVal sBasePath As String = "")

    Dim LI As HTMLLIElement
'FIXIT: Declare 'oChild' with an early-bound data type                                     FixIT90210ae-R1672-R1B8ZE
    Dim oChild As Object
    Dim p As HTMLParamElement
    Dim LIName As String
    Dim LILocal As String

    If sBasePath <> "" Then sBasePath = bdUnixDir(sBasePath, "")

    For Each LI In ULE.childNodes
        iStart = iStart + 1
        ReDim Preserve sAll(1, iStart) As String
        LIName = ""
        LILocal = ""

        For Each oChild In LI.childNodes

            Select Case oChild.nodeName
            Case "OBJECT"

                For Each p In oChild.childNodes
On Error Resume Next
                    If p.name = "Name" Then LIName = p.Value: LIName = p.Item  ' .Item

                    If p.name = "Local" Then LILocal = p.Value: LILocal = p.Value
                Next

                'If LILocal = "" Then LIName = LIName & "\"
                'LIName = bdUnixDir(sParent & LIName, "")
                iParent = iStart
                'If LILocal <> "" Then LILocal = bdUnixDir(LILocal)
                If sParent <> "" Then LIName = bdUnixDir(sParent, LIName)
                sAll(0, iStart) = LIName
                Debug.Print sAll(0, iStart)
                sAll(1, iStart) = sBasePath & LILocal
            Case "UL"
                LIName = bdUnixDir(LIName, "")
                sAll(0, iParent) = LIName
                getLI oChild, sAll, iStart, iParent, LIName, sBasePath
            End Select

        Next

    Next

End Sub

