VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAccessi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z


'===============================================
'       definizione variabili
'===============================================
Public lngHelpID As Long
Public intScheda As Integer
Public intAccesso As Integer
Public Sub ImpostaAccessi(intPAccesso As Integer)
    intAccesso = intPAccesso
End Sub


Public Sub Inizializza(lngPHelpID As Long, intPScheda As Integer, intPAccesso As Integer)
    lngHelpID = lngPHelpID
    intScheda = intPScheda
    intAccesso = intPAccesso
End Sub


