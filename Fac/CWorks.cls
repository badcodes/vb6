VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWorks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Id As Integer
Public DateTime As Date
Public Worker As CWorker
Public WorkRecords() As CWorkRecord
Public Property Get Salario() As Double
    On Error GoTo ErrorSalario
    Dim count As Integer
    count = UBound(WorkRecords)
    Dim i As Integer
    Dim sum As Double
    sum = 0
    For i = 0 To count
        sum = sum + WorkRecords(i).Salario
    Next
    Salario = sum
End Property
