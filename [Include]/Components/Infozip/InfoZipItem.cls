VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZipItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarFilename As String 'local copy
Private mvarSize As Long 'local copy
Private mvarPackedSize As Long 'local copy
Private mvarFactor As Long 'local copy
Private mvarMethod As String 'local copy
Private mvarCreatedDate As Date 'local copy
Private mvarCrc As Long 'local copy
Private mvarEncrypted As Boolean 'local copy
Private mvarFileType As VbFileAttribute
Friend Property Let FileType(ByVal vData As VbFileAttribute)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.FileType = Form1
    mvarFileType = vData
End Property


Public Property Get FileType() As VbFileAttribute
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileType
    FileType = mvarFileType
End Property



Friend Property Let Encrypted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Encrypted = 5
    mvarEncrypted = vData
End Property

Public Property Get Encrypted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Encrypted
    Encrypted = mvarEncrypted
End Property

Friend Property Let Crc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Crc = 5
    mvarCrc = vData
End Property


Public Property Get Crc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Crc
    Crc = mvarCrc
End Property

Friend Property Let CreatedDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreatedDate = 5
    mvarCreatedDate = vData
End Property

Public Property Get CreatedDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreatedDate
    CreatedDate = mvarCreatedDate
End Property

Friend Property Let Method(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Method = 5
    mvarMethod = vData
End Property

Public Property Get Method() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Method
    Method = mvarMethod
End Property

Friend Property Let Factor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Factor = 5
    mvarFactor = vData
End Property

Public Property Get Factor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Factor
    Factor = mvarFactor
End Property

Friend Property Let PackedSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PackedSize = 5
    mvarPackedSize = vData
End Property
Public Property Get PackedSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PackedSize
    PackedSize = mvarPackedSize
End Property

Friend Property Let Size(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Size = 5
    mvarSize = vData
End Property
Public Property Get Size() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Size
    Size = mvarSize
End Property

Friend Property Let Filename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Filename = 5
    mvarFilename = vData
End Property
Public Property Get Filename() As String
Attribute Filename.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Filename
    Filename = mvarFilename
End Property




