VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBookURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mfTextPdg As Boolean
Private mStep As Integer
Private mSolidPage(0 To 5) As String
Private mTextPdgs() As String
Private mIdxTextPdgs As Integer
Private mCountTextPdgs As Integer

Private mRangeIdx As Long
Private mStartPoint As Long
Private mRangePageFail As Boolean
Private mExt As String
Private mSkipCatalogPages As Boolean
Private mSkipContentPages As Boolean
Private mSkipStaticPages As Boolean

Public Property Get SkipStaticPages() As Boolean
    SkipStaticPages = mSkipStaticPages
End Property

Public Property Let SkipStaticPages(ByVal bValue As Boolean)
    mSkipStaticPages = bValue
End Property

Public Property Get SkipContentPage() As Boolean
    SkipContentPage = mSkipContentPages
End Property

Public Property Let SkipContentPage(ByVal bValue As Boolean)
    mSkipContentPages = bValue
End Property

Public Property Get SkipCatalogPages() As Boolean
    SkipCatalogPages = mSkipCatalogPages
End Property

Public Property Let SkipCatalogPages(ByVal bValue As Boolean)
    mSkipCatalogPages = bValue
End Property

Public Property Get Ext() As String
    Ext = mExt
End Property

Public Property Let Ext(ByVal StrValue As String)
    mExt = StrValue
End Property



Public Sub ReportFail(ByVal vUrl As String)
    Dim pStep As Long
    pStep = UrlToStep(vUrl)
    If pStep = mStep Then
        mRangeIdx = 0
        mStep = mStep + 1
    End If
    
End Sub

Public Sub SetStartPoint(ByVal vP As Long)
    mStartPoint = vP
End Sub
Public Sub SetTextPdg(ByVal vFilename As String)
    mIdxTextPdgs = 0
    mCountTextPdgs = MSSReader.SSLIB_ParseInfoRule(vFilename, mTextPdgs)
    If mCountTextPdgs > 0 Then mfTextPdg = True
End Sub

Private Function UrlToStep(ByRef vUrl As String) As Long
    Dim prv As String
    Dim ret As Long
    prv = LCase$(Left$(vUrl, 3))
    ret = -1
    Select Case prv
        Case "bac"
            ret = 6
        Case "att"
            ret = 7
        Case "cov"
            ret = 8
        Case "fow"
            ret = 9
        Case "bok"
            ret = 10
        Case "leg"
            ret = 11
        Case "ins"
            ret = 14
        Case Else
            prv = Left$(prv, 1)
            If prv = "!" Then ret = 12
            If prv = "0" Then ret = 13
    End Select
    UrlToStep = ret
End Function


Public Function GetNextUrl() As String
    Dim ret As String
    
    If mSkipStaticPages Then
        If mStep < 6 Then mStep = 6
    End If
    
    Select Case mStep
        Case 0
            ret = "bookinfo.dat"
        Case 1
            ret = "BookContents.dat"
        Case 2
            ret = "BookContents.xml"
        Case 3
            ret = "catalog.dat"
        Case 4
            ret = "Catalog.dat"
        Case 5
            ret = "InfoRule.dat"
    End Select
    If mfTextPdg = True And mStep > 5 Then
        If mCountTextPdgs > 0 And mIdxTextPdgs < mCountTextPdgs Then
            ret = mTextPdgs(mIdxTextPdgs)
            mIdxTextPdgs = mIdxTextPdgs + 1
        Else
            mStep = mStep + 1
        End If
    ElseIf mStep > 5 Then
        mRangeIdx = mRangeIdx + 1
        Select Case mStep
            Case 6
                ret = BuildRangeString("bac", 3, mExt, mRangeIdx)
            Case 7
                ret = BuildRangeString("att", 3, mExt, mRangeIdx)
            Case 8
                ret = BuildRangeString("cov", 3, mExt, mRangeIdx)
            Case 9
                ret = BuildRangeString("fow", 3, mExt, mRangeIdx)
            Case 10
                ret = BuildRangeString("bok", 3, mExt, mRangeIdx)
            Case 11
                ret = BuildRangeString("leg", 3, mExt, mRangeIdx)
            Case 12
                If Not mSkipCatalogPages Then
                    ret = BuildRangeString("!", 5, mExt, mRangeIdx)
                Else
                    mStep = mStep + 1
                    ret = GetNextUrl
                End If
            Case 13
                If Not mSkipContentPages Then
                    If mRangeIdx <= 1 Then mRangeIdx = mStartPoint
                    If mRangeIdx < 1 Then mRangeIdx = 1
                    ret = BuildRangeString("", 6, mExt, mRangeIdx)
                Else
                    mStep = mStep + 1
                    ret = GetNextUrl
                End If
            Case 14
                ret = BuildRangeString("ins", 3, mExt, mRangeIdx)
        End Select
    End If
    If mStep < 6 Then mStep = mStep + 1
    'Debug.Print "[" & ret & "]"
    GetNextUrl = ret
End Function

Private Function BuildRangeString(vPre As String, vNumLength As Long, vSuf As String, Optional vStart As Long = 1)
    Dim s As String
    Dim l As Long
    
    s = CStr(vStart)
    l = vNumLength - Len(s)
    If (l > 0) Then
        BuildRangeString = vPre & String$(l, "0") & s & vSuf
    Else
        BuildRangeString = vPre & s & vSuf
    End If

End Function

Private Sub Class_Initialize()
    mExt = ".pdg"
End Sub
