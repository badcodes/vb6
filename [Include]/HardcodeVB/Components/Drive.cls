VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'$ Uses DRIVE.BAS UTILITY.BAS

Enum EDriveType
    edtUnknown = 0
    edtNoRoot
    edtRemovable
    edtFixed
    edtRemote
    edtCDROM
    edtRAMDisk
End Enum

Public Enum EErrorDrive
    eeBaseDrive = 13020     ' CDrive
End Enum

Private sRoot As String
Private edtType As EDriveType
Private iTotalClusters As Long
Private iFreeClusters As Long
Private iSectors As Long
Private iBytes As Long
Private sLabel As String
Private iSerial As Long
Private fDriveMissing As Boolean

Private Sub Class_Initialize()
    InitAll
End Sub

Public Property Get FreeBytes() As Double
Attribute FreeBytes.VB_Description = "Free bytes available on the drive"
Attribute FreeBytes.VB_UserMemId = -502
    ' Always refresh size since free bytes might change
    GetSize
    If Not fDriveMissing Then
        FreeBytes = CDbl(iFreeClusters) * iSectors * iBytes
    End If
End Property

Public Property Get TotalBytes() As Double
    ' Get size info only on first access
    If iTotalClusters = 0 And Not fDriveMissing Then GetSize
    If Not fDriveMissing Then
        TotalBytes = CDbl(iTotalClusters) * iSectors * iBytes
    End If
End Property

Public Property Get Label() As String
    If Not fDriveMissing Then Label = sLabel
End Property

Public Property Get Serial() As String
    If Not fDriveMissing Then Serial = MUtility.FmtHex(iSerial, 8)
End Property

Public Property Get Kind() As EDriveType
    Kind = edtType
End Property

Public Property Get KindStr() As String
    KindStr = Choose(edtType + 1, "Unknown", "Invalid", "Floppy", _
                                  "Fixed", "Network", "CD-ROM", "RAM")
    If fDriveMissing Then KindStr = KindStr & " Missing"
End Property

Public Property Get Number() As Integer
    Number = Asc(sRoot) - Asc("A") + 1
    ' Network drives are zero
    If Number > 26 Then Number = 0
End Property

Public Property Get Root() As Variant
Attribute Root.VB_UserMemId = 0
    Root = sRoot
End Property

Public Property Let Root(vRootA As Variant)
    ' Some properties won't work for \\server\share\ drives on Windows 95
    sRoot = UCase(vRootA)  ' Convert to string
    InitAll
End Property

Private Sub InitAll()
    sLabel = sEmpty: iSerial = 0
    iSectors = 0: iBytes = 0: iFreeClusters = 0: iTotalClusters = 0
    fDriveMissing = False
    ' Empty means get current drive
    If sRoot = sEmpty Then sRoot = Left$(CurDir$, 3)
    ' Get drive type ordinal
    edtType = GetDriveType(sRoot)
    ' If invalid root string, try it with terminating backslash
    If edtType = edtNoRoot Then edtType = GetDriveType(sRoot & "\")
    Select Case edtType
    Case edtUnknown, edtNoRoot
        Dim iDrive As String
        iDrive = Val(sRoot)
        If iDrive >= 1 And iDrive <= 26 Then
            sRoot = Chr$(iDrive + Asc("A") - 1) & ":\"
        Else
            sRoot = sEmpty
        End If
        ' Start over
        InitAll
    Case edtRemovable, edtFixed, edtRemote, edtCDROM, edtRAMDisk
        ' If you got here, drive is valid, but root might not be
        If Right$(sRoot, 1) <> "\" Then sRoot = sRoot & "\"
        GetLabelSerial
    Case Else ' Shouldn't happen
        BugAssert True
    End Select
End Sub

Public Property Let Label(sLabelA As String)
    If SetVolumeLabel(sRoot, sLabelA) Then sLabel = sLabelA
End Property

Private Sub GetLabelSerial()
    sLabel = String$(cMaxPath, 0)
    Dim afFlags As Long, iMaxComp As Long
    Call GetVolumeInformation(sRoot, sLabel, cMaxPath, iSerial, _
                              iMaxComp, afFlags, sNullStr, 0)
    fDriveMissing = Err.LastDllError
    sLabel = MUtility.StrZToStr(sLabel)
End Sub

Private Sub GetSize()
    Call GetDiskFreeSpace(sRoot, iSectors, iBytes, _
                          iFreeClusters, iTotalClusters)
    fDriveMissing = Err.LastDllError
End Sub
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Drive"
        Select Case e
        Case eeBaseDrive
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

