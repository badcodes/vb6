VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnumVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Implement VB version of IEnumVARIANT (from type library interface)
Implements IVBEnumVARIANT

Public Enum EErrorEnumVariant
    eeBaseEnumVariant = 13050   ' CEnumVariant
End Enum

Private connect As IVariantWalker

' Delegators must connect themselves so we can call back
Sub Attach(connectA As IVariantWalker)
    Set connect = connectA
End Sub

Private Sub Class_Initialize()
    ' Only executes once for life of program
    If MEnumVariant.fNotFirstTime = False Then
        MEnumVariant.fNotFirstTime = True
        ' There's only one v-table for the object, so modify it once
        Dim iev As IVBEnumVARIANT
        Set iev = Me
        ' Ignore item 1: QueryInterface
        ' Ignore item 2: AddRef
        ' Ignore item 3: Release
        ReplaceVtableEntry ObjPtr(iev), 4, AddressOf MEnumVariant.BasNext
        ReplaceVtableEntry ObjPtr(iev), 5, AddressOf MEnumVariant.BasSkip
        ' Ignore item 6: Reset
        ' Ignore item 7: Clone
    End If
    BugLocalMessage "CEnumVariant:Initialize"
End Sub

Private Sub Class_Terminate()
    BugLocalMessage "CEnumVariant:Terminate"
End Sub

'' Dummy versions of implemented functions are replaced by standard module versions

Private Sub IVBEnumVARIANT_Next(ByVal cv As Long, v As Variant, ByVal cvFetched As Long)
    BugLocalMessage "Dummy Next"
End Sub

Private Sub IVBEnumVARIANT_Skip(ByVal cv As Long)
    BugLocalMessage "Dummy Skip"
End Sub

' Can be implemented directly--no need to mess with the v-table
Private Sub IVBEnumVARIANT_Reset()
    BugAssert Not connect Is Nothing
    connect.Reset
    BugLocalMessage "Reset"
End Sub

' Not implemented--just raise an error
Private Sub IVBEnumVARIANT_Clone(ppenum As stdole.IEnumVARIANT)
    Err.Raise &H80004001     ' E_NOTIMPL
End Sub

'' Object methods called by standard module v-table functions contain implementation

Function ClsNext(v As Variant) As Boolean
    BugAssert Not connect Is Nothing
    ClsNext = connect.More(v)
End Function

Sub ClsSkip(c As Long)
    BugAssert Not connect Is Nothing
    connect.Skip c
    BugLocalMessage "Skip"
End Sub

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".EnumVariant"
        Select Case e
        Case eeBaseEnumVariant
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

