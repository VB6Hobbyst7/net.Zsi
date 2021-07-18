VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGestMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarCaption As String 'Copia locale.
Private mvarHelpContextID As Long 'Copia locale.
Private mvarHasSeparator As Boolean 'Copia locale.
Private mvarKey As Variant 'Copia locale.
Private mvarParentKey As Variant 'Copia locale.
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarModulo As String 'Copia locale.
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarMenuID As Long 'Copia locale.
Public Property Let MenuID(ByVal vData As Long)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.MenuID = 5
    mvarMenuID = vData
End Property


Public Property Get MenuID() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.MenuID
    MenuID = mvarMenuID
End Property



Public Property Let Modulo(ByVal vData As String)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.Modulo = 5
    mvarModulo = vData
End Property


Public Property Get Modulo() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Modulo
    Modulo = mvarModulo
End Property



Public Property Let ParentKey(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.ParentKey = 5
    mvarParentKey = vData
End Property


Public Property Get ParentKey() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.ParentKey
    ParentKey = mvarParentKey
End Property



Public Property Let key(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Get key() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Key
    key = mvarKey
End Property



Public Property Let HasSeparator(ByVal vData As Boolean)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.HasSeparator = 5
    mvarHasSeparator = vData
End Property


Public Property Get HasSeparator() As Boolean
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.HasSeparator
    HasSeparator = mvarHasSeparator
End Property



Public Property Let HelpContextID(ByVal vData As Long)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.HelpContextID = 5
    mvarHelpContextID = vData
End Property


Public Property Get HelpContextID() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.HelpContextID
    HelpContextID = mvarHelpContextID
End Property



Public Property Let Caption(ByVal vData As String)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Public Property Get Caption() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property



