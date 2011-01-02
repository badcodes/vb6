VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GInfozipShared"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function strConvX(srcString As Variant, Conversion As VbStrConv, Optional LCID As LOCALEID) As String
    strConvX = MInfoZipShared.strConvX(srcString, Conversion, LCID)
End Function

Public Function CBytesToStr(ByRef CBytes() As Byte) As String
    CBytesToStr = MInfoZipShared.CBytesToStr(CBytes())
End Function

    
    
    
    

Public Sub StrToCBytes(ByVal strUnicode As String, ByRef CBytes() As Byte)
   MInfoZipShared.StrToCBytes strUnicode, CBytes()
End Sub

    
    
    
    
    
       

Public Function CleanZipFilename(sInCome) As String
    ' Can't translate
End Function






