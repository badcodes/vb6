VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ReaderHandlers() As IReaderHandler
Private DocumentHandlers() As IDocumentHandler
Private ViewerHandlers() As IViewerHandler
Private cReaderHND As Integer
Private cDocHND As Integer
Private cViewerHND As Integer


Private Function SelectHandlerForDocument(ByRef vFilename As String) As IDocumentHandler
Dim i As Integer
For i = 0 To cDocHND - 1
    If DocumentHandlers(i).canhandler(vFilename) Then
        Set SelectHandlerForDocument = DocumentHandlers(i)
    End If
Next
Set SelectHandlerForDocument = Nothing
End Function

Public Function OpenFile(ByRef vFilename As String) As CDocument
    Dim hnd As IDocumentHandler
    Set hnd = SelectHandlerForDocument(vFilename)
    If hnd Is Nothing Then
        Set OpenFile = Nothing
    Else
        Set OpenFile = hnd.OpenFile(vFilename)
        Set OpenFile.Handler = hnd
    End If
End Function

Public Function CloseDocument(ByRef vDocument As CDocument) As Boolean
    Set vDocument.Handler = Nothing
    Set vDocument = Nothing
End Function

Public Function OpenDocument(ByRef vDocument As CDocument, ByRef vEntry As String) As String
    OpenDocument = vDocument.OpenEntry(vEntry)
End Function

Public Function OpenHistory(ByRef vHistory As String) As CDocument
    Dim Filename As String
    Dim Entry As String
    
End Function

Public Function SaveHistory(ByRef vDocument As CDocument, ByRef vEntry As String) As String

End Function

Public Function BeforeViewing(ByRef vDocument As CDocument, _
    ByRef vUrl As String, ByRef vParam As Object) As Boolean
    
End Function

Public Sub AfterViewed(ByRef vDocument As CDocument, _
    ByRef vUrl As String, ByRef vCancel As Boolean, _
    ByRef vParam As Object)
    
End Sub


Public Function StatusChanged(ByRef vText As String) As Boolean

End Function

Public Function Init() As Boolean

End Function
