VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBookTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event TestCompleted()
Public Event StatusChange(ByRef vText As String)
Private mIsValid As Boolean
Private mIsTextPdg As Boolean
Private mStartPage As Long

Private mHttp As IWinHttp
Attribute mHttp.VB_VarHelpID = -1
Private WithEvents mRealHttp As CWinHttpSimple
Attribute mRealHttp.VB_VarHelpID = -1

Private mStep As Integer
Private mHeader As String
Private mRootUrl As String
Private mFilename As String
Private mDirectory As String

Private mPageCount As Long
Private mPageStart As Long
Private mIsJPGBook As Boolean
Private mJPGRootUrl As String
Private mState(1 To 4) As Boolean

Public Property Get IsJpgBook() As Boolean
    IsJpgBook = mIsJPGBook
End Property

Public Property Get StartPage() As Long
    StartPage = mStartPage
End Property

Public Property Get IsTextPdg() As Boolean
    IsTextPdg = mIsTextPdg
End Property

Public Property Get IsValid() As Boolean
    IsValid = mIsValid
End Property

Private Sub Reset()
    mStep = 0
    mIsValid = False
    mIsTextPdg = False
    mIsJPGBook = False
    mStartPage = 0
End Sub

Public Sub Test(ByRef vTask As CTask)
    Reset
    
    mIsJPGBook = vTask.IsJpgBook
    
    mHeader = vTask.bookInfo(SSF_HEADER)

    mRootUrl = vTask.bookInfo(SSF_URL)
    mJPGRootUrl = vTask.bookInfo(SSF_JPGURL)
    
    mDirectory = BuildPath(vTask.Directory)
    mPageCount = StringToLong(vTask.bookInfo(SSF_PagesCount))
    If mPageCount < 1 Then mPageCount = 9999
    mPageStart = StringToLong(vTask.bookInfo(SSF_StartPage))
    If mPageStart < 1 Then mPageStart = 1
    
    StartTest
End Sub
Private Sub StartTest()
    mStep = 1
    mFilename = "InfoRule.dat"
    testFile
End Sub
Private Sub Class_Initialize()
    Reset
    'Set mHttp = New CWinHTTP
End Sub

Private Sub testFile()
'    Dim pLocal As String
'    pLocal = BuildPath(mDirectory, mFilename)
'    If FileExists(pLocal) Then
'        Report 200
'    Else
'    If mIsJPGBook Then
'        mIsValid = True
'        RaiseEvent TestCompleted
'        Exit Sub
'    End If
    If FileExists(mDirectory & mFilename) Then
        RaiseEvent StatusChange(mFilename & " ´æÔÚ£¬ÂÔ¹ý²âÊÔ")
        Report 200
        Exit Sub
    End If
        If mRealHttp Is Nothing Then Set mRealHttp = New CWinHttpSimple
        Set mHttp = mRealHttp
        With mHttp
            .Init
            .Method = "GET"
            .URL = BuildPath(mRootUrl, mFilename, lnpsUnix)
            .Header = mHeader
            .TestMode = True
            .OpenConnect True
            .Send
        End With
'        RaiseEvent StatusChange(mFilename & " " & mHttp.StatusText)
'        Report mHttp.Status
'    End If
End Sub
'Private Sub mRealHttp_OnTestCompleted(ByVal vTag As Variant, ByVal vStatus As Long, ByVal vText As String)
'    'Set mRealHttp = Nothing
'    RaiseEvent StatusChange(mFilename & " " & vText)
'    Report vStatus
'End Sub
Private Sub Report(vStatus As Long)
    If mStep = 1 Then
        If vStatus = 200 Or vStatus = 404 Then
            mState(1) = True
            mIsValid = True
        Else
            mState(1) = False
            mIsValid = False
        End If
        If vStatus = 200 Then
            mIsTextPdg = True
        Else
            mIsTextPdg = False
        End If
        If mIsValid = False Or mIsTextPdg = True Then
            GoTo TestCompleted
        Else
            mStep = 2
            mFilename = "000001.pdg"
            testFile
        End If
    ElseIf mStep = 2 Then
        If vStatus = 200 Then
            mStartPage = 1
            GoTo TestCompleted
        Else
            mStep = 3
            mFilename = BuildRangeString("", 6, ".pdg", mPageStart)
            testFile
        End If
    ElseIf mStep = 3 Then
        If vStatus = 200 Then
            mStartPage = mPageStart
            GoTo TestCompleted
        Else
            mPageStart = mPageStart + 1
            If mPageStart > mPageCount Then
                mStartPage = -1
                GoTo TestCompleted
            End If
            mFilename = BuildRangeString("", 6, ".pdg", mPageStart)
            testFile
        End If
    End If
    Exit Sub
TestCompleted:
    If mJPGRootUrl <> "" And mIsJPGBook Then
        mIsValid = True
        RaiseEvent TestCompleted
    Else
        RaiseEvent TestCompleted
    End If
    Exit Sub

End Sub
'Private Sub mHttp_OnError(ByVal vTag As Variant, ByVal ErrorNum As Long, ByVal ErrorDescription As String)
'    Report ErrorNum
'End Sub

'
'Private Sub mHttp_OnResponseFinished(ByVal vTag As Variant)
'
'End Sub
'
'Private Sub mHttp_OnResponseStart(ByVal vTag As Variant, ByVal Status As Long)
'    Report Status
'    Debug.Print Status
'End Sub
'Private Property Get StatusText() As String
'    StatusText = mStatusText
'End Property



Private Sub mRealHttp_OnTestCompleted(ByVal vTag As Variant, ByVal vStatus As Long, ByVal vText As String)
        RaiseEvent StatusChange(mFilename & " " & vText)
        Report vStatus
End Sub
