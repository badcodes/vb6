VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum EAppearance
    eaFlat
    ea3D
End Enum

Enum ESortModeList
    esmlUnsorted
    esmlSortVal
    esmlSortText
    esmlSortBin
    esmlSortLen
    esmlShuffle
End Enum

Enum ESortError
    eseNone = 1000
    eseItemNotFound
    eseOutOfRange
    eseDuplicateNotAllowed
End Enum


