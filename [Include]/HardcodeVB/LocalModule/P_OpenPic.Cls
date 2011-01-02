VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COpenPictureFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EErrorOpenPictureFile
    eeBaseOpenPictureFile = 13680   ' COpenPictureFile
End Enum

' Bit flags for filter types
Public Enum EFilterPicture
    efpAllPicture = &H1
    efpBitmap = &H2
    efpMetafile = &H4
    efpIcon = &H8
    efpCursor = &H10
    efpWave = &H20
    efpAllFile = &H40
    efpEverything = &H7F
End Enum
                       
Private frm As New FOpenPictureFile

Sub Load(Optional ByVal Left As Single = -1#, _
         Optional ByVal Top As Single = -1#)
    frm.Dummy = False
    If Left <> -1# Then frm.Left = Left
    If Top <> -1# Then frm.Top = Top
    frm.Show vbModal
End Sub

Property Get FileTitle() As String
    FileTitle = frm.FileTitle
End Property

Property Get FileName() As String
    FileName = frm.FileName
End Property

Property Let FileName(sFilePathA As String)
    frm.FileName = sFilePathA
End Property

Property Get InitDir() As String
    InitDir = frm.InitDir
End Property

Property Let InitDir(sInitDirA As String)
    frm.InitDir = sInitDirA
End Property

Property Get PicType() As Integer
    PicType = frm.PicType
End Property

Property Get FilterType() As Long
    FilterType = frm.FilterType
End Property

Property Let FilterType(afFilterA As Long)
    frm.FilterType = afFilterA
End Property




