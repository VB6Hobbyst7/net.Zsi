VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Dim mvarKey As String
Dim mvarValore As String

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let Key(ByVal v As String)
    mvarKey = v
End Property

Public Property Get Valore() As Double
    Valore = mvarValore
End Property

Public Property Let Valore(ByVal v As Double)
    mvarValore = v
End Property

