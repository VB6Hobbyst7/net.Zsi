VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe wrapper verso .NET per esporre le Collection VB6 come Object

Dim mColCollection As VBA.Collection

Public Property Get CollectionWrapper() As Object
    Set CollectionWrapper = mColCollection
End Property

Public Property Set CollectionWrapper(vdata As Object)
    Set mColCollection = vdata
End Property

Private Sub Class_Initialize()
    Set mColCollection = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Set mColCollection = Nothing
End Sub

