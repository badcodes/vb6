Attribute VB_Name = "MAllocator"
Option Explicit

Private alloc As IVBMalloc

Property Get Allocator() As IVBMalloc
    If alloc Is Nothing Then SHGetMalloc alloc
    Set Allocator = alloc
End Property

