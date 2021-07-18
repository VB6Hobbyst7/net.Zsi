VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvarKey As String 'Copia locale.
Private mvarStato As String 'Copia locale.
Private mvarItemData As Variant

Public Property Let Stato(ByVal vData As String)
    mvarStato = vData
End Property


Public Property Get Stato() As String
    Stato = mvarStato
End Property


Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property


Public Property Get Key() As String
    Key = mvarKey
End Property


Public Property Let ItemData(ByVal vData As Variant)
    mvarItemData = vData
End Property


Public Property Get ItemData() As Variant
    ItemData = mvarItemData
End Property

