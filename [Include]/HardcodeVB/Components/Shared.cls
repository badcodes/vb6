VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GShared"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Put public properties that delegate shared data here

Property Get System() As CSystem
    Set System = MShared.System
End Property

Property Get Video() As CVideo
    Set Video = MShared.Video
End Property

Property Get Keyboard() As CKeyboard
    Set Keyboard = MShared.Keyboard
End Property


