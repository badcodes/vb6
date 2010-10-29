VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMRUFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_lMaxCount As Long
Private m_lFileCount As Long
Private m_sFiles() As String

Public Property Get MaxFileCount() As Long
    MaxFileCount = m_lMaxCount
End Property
Public Property Let MaxFileCount(ByVal lMax As Long)
    If (m_lMaxCount <> lMax) Then
        ReDim Preserve m_sFiles(1 To lMax) As String
        m_lMaxCount = lMax
    End If
End Property
Public Property Get FileCount() As Long
    FileCount = m_lFileCount
End Property
Public Property Get file(ByVal iIndex As Long) As String
    file = m_sFiles(iIndex)
End Property
Public Property Get FileExists(ByVal iIndex As Long) As Boolean
    FileExists = (m_sFiles(iIndex) <> "")
End Property
Public Property Get MenuCaption(ByVal iIndex As Long) As String
Dim lLen As Long
Dim sOut As String
Dim iPos As Long
Dim iLastPos As Long

    sOut = m_sFiles(iIndex)
    ' Easier way using API probably, but here goes:
    lLen = Len(m_sFiles(iIndex))
    If (lLen > 32) Then
        For iPos = lLen To 1 Step -1
            If (Mid$(m_sFiles(iIndex), iPos, 1) = "\") Then
                If (lLen - iPos) > 32 Then
                    If (iLastPos = 0) Then
                        sOut = "..." & Mid$(m_sFiles(iIndex), iPos)
                    Else
                        sOut = "..." & Mid$(m_sFiles(iIndex), iLastPos)
                    End If
                Else
                    iLastPos = iPos
                End If
            End If
        Next iPos
    End If
    MenuCaption = "&" & iIndex & ")  " & sOut
    
End Property
Public Sub AddFile(ByVal sfileName As String)
Dim iFile As Long
Dim iExistIndex As Long
Dim iSwap As Long
    ' does the file already exist?
    For iFile = 1 To m_lFileCount
        If (UCase$(m_sFiles(iFile)) = UCase$(sfileName)) Then
            iExistIndex = iFile
            Exit For
        End If
    Next iFile
    
    ' If it exists swap all the entries above it down
    ' by one, otherwise swap all entries down by one:
    If (iExistIndex = 0) Then
        If (m_lFileCount < MaxFileCount) Then
            m_lFileCount = m_lFileCount + 1
        End If
        iSwap = m_lFileCount
    Else
        iSwap = iExistIndex
    End If
    
    For iFile = iSwap - 1 To 1 Step -1
        m_sFiles(iFile + 1) = m_sFiles(iFile)
    Next iFile
    
    ' Add this entry:
    m_sFiles(1) = sfileName
    
End Sub
Public Function Save(ByRef cR As cRegistry) As Boolean
On Error GoTo SaveError

Dim iFile As Long

    cR.ValueKey = "MaxCount"
    cR.ValueType = REG_DWORD
    cR.Value = MaxFileCount

    cR.ValueKey = "Count"
    cR.ValueType = REG_DWORD
    cR.Value = m_lFileCount
    
    For iFile = 1 To m_lFileCount
        cR.ValueKey = "File" & iFile
        cR.ValueType = REG_SZ
        cR.Value = m_sFiles(iFile)
    Next iFile
    
    Save = True
    Exit Function
SaveError:
    Exit Function
End Function
Public Function Load(ByRef cR As cRegistry) As Boolean
On Error GoTo LoadError
    
Dim iFile As Long

    cR.ValueKey = "MaxCount"
    If (cR.KeyExists) Then
        cR.ValueType = REG_DWORD
        MaxFileCount = cR.Value
        
        If (MaxFileCount > 0) Then
            cR.ValueKey = "Count"
            cR.ValueType = REG_DWORD
            m_lFileCount = cR.Value
            
            For iFile = 1 To m_lFileCount
                cR.ValueKey = "File" & iFile
                cR.ValueType = REG_SZ
                m_sFiles(iFile) = cR.Value
            Next iFile
            
        End If
    End If
    Load = True
    Exit Function
LoadError:
    Exit Function
End Function


