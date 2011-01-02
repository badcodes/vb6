VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComplexDataConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' TODO: declare a recordset variable to be the datasource for this class.  For example:

'Private WithEvents RS As ADODB.RecordSet

Public Property Get DataSource() As DataSource
    ' TODO: return the datasource object for this class.  For example:
   
    'Set DataSource = RS.DataSource
End Property
Public Property Set DataSource(ByVal objDataSource As DataSource)
    '  TODO: set the datasource object for this class.  For example:
    
    'Set RS.DataSource = objDataSource
    'PropertyChanged "DataSource"
End Property
Public Property Get DataMember() As DataMember
    ' TODO: return the DataMember.  For example:
    
    'Set DataMember = RS.DataMember
End Property
Public Property Let DataMember(ByVal DataMember As DataMember)
    ' TODO: assign a new DataMember.  For example:
    
    'Set RS.DataMember = DataMember
    'PropertyChanged "DataMember"
End Property

