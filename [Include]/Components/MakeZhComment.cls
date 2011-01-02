VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMakeZhComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VB class for generate zhComment file from a list of files."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvartitle As String 'local copy
Private mvarsaveTo As String 'local copy
Private mvarauthor As String 'local copy
Private mvarpublisher As String 'local copy
Private mvarcatalog As String 'local copy
Private mvardefaultFile As String 'local copy
Private mvarhhcFile As String 'local copy
Private mvarcontentFile As String 'local copy
Private mvarshowLeft As Boolean 'local copy
Private mvarshowMenu As Boolean 'local copy
Private mvarshowStatusBar As Boolean 'local copy
Private mvarfileDate As String 'local copy

Private d_Publisher As String
Private d_HHCFile As String
Private d_Date As String
Private d_Saveto As String

Public Property Let fileDate(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.fileDate = 5
    mvarfileDate = vData
End Property

Public Property Get fileDate() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.fileDate
    fileDate = mvarfileDate
End Property

Public Property Let showStatusBar(ByVal vData As Boolean)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.showStatusBar = 5
    mvarshowStatusBar = vData
End Property

Public Property Get showStatusBar() As Boolean
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.showStatusBar
    showStatusBar = mvarshowStatusBar
End Property

Public Property Let showMenu(ByVal vData As Boolean)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.showMenu = 5
    mvarshowMenu = vData
End Property

Public Property Get showMenu() As Boolean
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.showMenu
    showMenu = mvarshowMenu
End Property

Public Property Let showLeft(ByVal vData As Boolean)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.showLeft = 5
    mvarshowLeft = vData
End Property

Public Property Get showLeft() As Boolean
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.showLeft
    showLeft = mvarshowLeft
End Property

Public Property Let contentFile(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.contentFile = 5
    mvarcontentFile = vData
End Property

Public Property Get contentFile() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.contentFile
    contentFile = mvarcontentFile
End Property

Public Property Let hhcFile(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.hhcFile = 5
    mvarhhcFile = vData
End Property

Public Property Get hhcFile() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.hhcFile
    hhcFile = mvarhhcFile
End Property

Public Property Let defaultFile(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.defaultFile = 5
    mvardefaultFile = vData
End Property

Public Property Get defaultFile() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.defaultFile
    defaultFile = mvardefaultFile
End Property

Public Property Let catalog(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.catalog = 5
    mvarcatalog = vData
End Property

Public Property Get catalog() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.catalog
    catalog = mvarcatalog
End Property

Public Property Let publisher(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.publisher = 5
    mvarpublisher = vData
End Property

Public Property Get publisher() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.publisher
    publisher = mvarpublisher
End Property

Public Property Let author(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.author = 5
    mvarauthor = vData
End Property

Public Property Get author() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.author
    author = mvarauthor
End Property

Public Function make(FileList() As String) As String
    Dim iStart As Long
    Dim iEnd As Long
    iStart = LBound(FileList)
    iEnd = UBound(FileList)
    If iEnd - iStart < 0 Then Exit Function
    
    
    If saveTo = "" Then saveTo = d_Saveto
    If linvblib.FileExists(saveTo) = False Then Exit Function
    If publisher = "" Then publisher = d_Publisher
    If fileDate = "" Then fileDate = d_Date
    If title = "" Then title = linvblib.GetFileName(FileList(iStart))
    
    
    Dim lfor As Long
    Dim sFilename As String
    Dim sExtname As String
    Dim bRequireDefaultFile As Boolean
    Dim bRequireHHC As Boolean
    Dim sDefaultfile As String
    Dim iCountSL As Integer ' Count "\" in sFilename
    Dim iMinSL As Integer
    iMinSL = 100 ' 设为最大值

    With Me
        If .defaultFile = "" Then bRequireDefaultFile = True
        If .hhcFile = "" And .contentFile = "" Then bRequireHHC = True
    End With

    For lfor = iStart To iEnd
        If bRequireDefaultFile Or bRequireHHC Then
            sFilename = FileList(lfor)
            sExtname = LCase$(linvblib.GetExtensionName(sFilename))
            If bRequireHHC And sExtname = "hhc" Then
                Me.hhcFile = sFilename
                bRequireHHC = False
            ElseIf (sExtname = "html" Or sExtname = "htm") And _
               bRequireDefaultFile And _
               IsWebsiteDefaultFile(sFilename) Then
                iCountSL = linvblib.charCountInStr(sFilename, "\", vbBinaryCompare)
                If iCountSL < iMinSL Then
                    iMinSL = iCountSL
                    sDefaultfile = sFilename
                    If iMinSL = 0 Then bRequireDefaultFile = False
                End If
            End If
        End If
    Next

    If sDefaultfile = "" Then sDefaultfile = FileList(iStart)
    Me.defaultFile = sDefaultfile
    If Me.hhcFile = "" Then Me.hhcFile = d_HHCFile
    Dim fNUM As Integer
    fNUM = FreeFile
    Open Me.saveTo For Output As fNUM
    Print #fNUM, "[Info]"
    Print #fNUM, "Title=" & title
    Print #fNUM, "Author=" & author
    Print #fNUM, "Date=" & fileDate
    Print #fNUM, "Publisher=" & publisher
    Print #fNUM, "Catalog=" & catalog
    Print #fNUM, "DefaultFile=" & defaultFile
    Print #fNUM, "HHCfile=" & hhcFile
    Print #fNUM, "ContentFile=" & contentFile
    Print #fNUM, "[Style]"
    Print #fNUM, "ShowLeft=" & CStr(Abs(CInt(showLeft)))
    Print #fNUM, "ShowMenu=0" & CStr(Abs(CInt(showMenu)))
    Print #fNUM, "ShowStatusBar=" & CStr(Abs(CInt(showStatusBar)))
    Close fNUM
    make = saveTo
End Function

Public Property Let saveTo(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.saveTo = 5
    mvarsaveTo = vData
End Property

Public Property Get saveTo() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.saveTo
    saveTo = mvarsaveTo
End Property

Public Property Let title(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.title = 5
    mvartitle = vData
End Property

Public Property Get title() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.title
    title = mvartitle
End Property

Private Sub Class_Initialize()
    d_Publisher = Environ$("username")
    d_Date = Date$
    d_HHCFile = "none"
    d_Saveto = Environ$("temp") & "\zh.ini"
End Sub

     

