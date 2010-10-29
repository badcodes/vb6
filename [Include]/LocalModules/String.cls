VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mString As String
Private m_lLength As Long
Private mvarCompareMethod As VbCompareMethod

Public Enum IfNotFound
    ReturnOriginal = 1
    ReturnEmpty = 0
End Enum

Public Property Let CompareMethod(ByVal vData As VbCompareMethod)

    mvarCompareMethod = vData

End Property

Public Property Get CompareMethod() As VbCompareMethod

    CompareMethod = mvarCompareMethod

End Property

Public Function StartsWith(sSubString As String) As Boolean

    If InStr(1, mString, sSubString, mvarCompareMethod) = 1 Then StartsWith = True

End Function

Public Function EndsWith(sSubString As String) As Boolean

    If InStrRev(mString, sSubString, , mvarCompareMethod) = m_lLength - Len(sSubString) + 1 Then EndsWith = True

End Function

Public Function Remove(Optional iStartIndex As Long = 1, Optional iCount As Long = -1) As String

    If iStartIndex < 0 Then iStartIndex = 1

    If iStartIndex > m_lLength Then Exit Function

    If iCount < 0 Then iCount = m_lLength - iStartIndex + 1

    If iCount > m_lLength - iStartIndex + 1 Then iCount = m_lLength - iStartIndex + 1
    Remove = Left$(mString, iStartIndex - 1) & Right$(mString, m_lLength - iStartIndex - iCount + 1)
    mString = Remove
    Call StringChange

End Function

Public Function IndexOf(sSubString As String, Optional iStartIndex As Long = 1) As Long

    If iStartIndex > m_lLength Then Exit Function
    IndexOf = InStr(iStartIndex, mString, sSubString, mvarCompareMethod)

End Function

Public Function LastIndexOf(sSubString As String, Optional iStartIndex As Long = -1) As Long

    If iStartIndex > m_lLength Then Exit Function
    LastIndexOf = InStrRev(mString, sSubString, iStartIndex, mvarCompareMethod)

End Function

Public Function SubString(Optional iStartIndex As Long = 1, Optional iCount As Long = -1) As String

    If iStartIndex < 0 Then iStartIndex = 1

    If iStartIndex > m_lLength Then Exit Function

    If iCount < 0 Then iCount = m_lLength - iStartIndex + 1

    If iCount > m_lLength - iStartIndex + 1 Then iCount = m_lLength - iStartIndex + 1
    SubString = Mid$(mString, iStartIndex, iCount)

End Function

Public Property Get Length() As Long

    Length = m_lLength

End Property

Public Property Get Value() As String
Attribute Value.VB_Description = "String Value (Default)"
Attribute Value.VB_UserMemId = 0

    Value = mString

End Property

Public Property Let Value(sValue As String)

    mString = sValue
    Call StringChange

End Property

Private Sub Class_Initialize()

    mvarCompareMethod = vbBinaryCompare

End Sub

Public Sub Create(sValue As String)

    Me.Value = sValue

End Sub

Private Sub StringChange()

    m_lLength = Len(mString)

End Sub

Public Function CountOf(strSmall As String) As Long

    Dim lenSmall As Long
    Dim posStart As Long, nextPos As Long
    lenSmall = Len(strSmall)

    If m_lLength < lenSmall Or lenSmall = 0 Then Exit Function
    posStart = InStr(1, mString, strSmall, mvarCompareMethod)

    Do Until posStart < 1
        CountOf = CountOf + 1
        nextPos = posStart + lenSmall

        If nextPos > m_lLength Then Exit Do
        posStart = InStr(nextPos, mString, strSmall, mvarCompareMethod)
    Loop

End Function

' Search from end to beginning and return the right side of the string
Public Function RightRight(RFind As String, Optional RetError As IfNotFound = ReturnOriginal) As String

    Dim K As Long
    K = InStrRev(mString, RFind, , mvarCompareMethod)

    If K = 0 Then
        RightRight = IIf(RetError = ReturnOriginal, mString, "")
    Else
        RightRight = Mid$(mString, K + 1, m_lLength)
    End If

End Function

' Search from the beginning to end and return the left size of the string
Public Function LeftLeft(LFind As String, Optional RetError As IfNotFound = ReturnOriginal) As String

    Dim K As Long
    K = InStr(1, mString, LFind, mvarCompareMethod)

    If K = 0 Then
        LeftLeft = IIf(RetError = ReturnOriginal, mString, "")
    Else
        LeftLeft = Left$(mString, K - 1)
    End If

End Function

' Search from the beginning to end and return the right size of the string
Public Function LeftRight(LFind As String, Optional RetError As IfNotFound = ReturnOriginal) As String

    Dim K As Long
    K = InStr(1, mString, LFind, mvarCompareMethod)

    If K = 0 Then
        LeftRight = IIf(RetError = ReturnOriginal, mString, "")
    Else
        LeftRight = Right$(mString, (m_lLength - Len(LFind)) - K + 1)
    End If

End Function

' Search from the beginning to end and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function LeftRange(strFrom As String, strTo As String, Optional RetError As IfStringNotFound = ReturnOriginal) As String

    Dim K As Long, Q As Long
    K = InStr(1, mString, strFrom, mvarCompareMethod)

    If K > 0 Then
        Q = InStr(K + Len(strFrom), mString, strTo, mvarCompareMethod)

        If Q > K Then
            LeftRange = Mid$(mString, K + Len(strFrom), (Q - K) - Len(strFrom))
        Else
            LeftRange = IIf(RetError = ReturnOriginal, mString, "")
        End If

    Else
        LeftRange = IIf(RetError = ReturnOriginal, mString, "")
    End If

End Function

' Search from the end to beginning and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function RightRange(strFrom As String, strTo As String, Optional RetError As IfStringNotFound = ReturnOriginal) As String

    Dim K As Long, Q As Long
    K = InStrRev(mString, strTo, , mvarCompareMethod)

    If K > 0 Then
        Q = InStrRev(mString, strFrom, K, mvarCompareMethod)

        If Q > 0 Then
            RightRange = Mid$(mString, Q + Len(strFrom), (K - Q) - Len(strFrom))
        Else
            RightRange = IIf(RetError = ReturnOriginal, mString, "")
        End If

    Else
        RightRange = IIf(RetError = ReturnOriginal, mString, "")
    End If

End Function

