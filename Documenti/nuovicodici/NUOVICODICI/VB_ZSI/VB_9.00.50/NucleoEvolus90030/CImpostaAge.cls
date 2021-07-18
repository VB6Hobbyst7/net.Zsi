VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImpostaAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'===============================================
'       definizione variabili
'===============================================
Public lngHelpID As Long 'help ID della form
'Public strNomeCtrl As String 'nome controllo
'Public setEvento As enmAgtEventi 'evento da gestire

Public strAgente As String 'agente da eseguire
Public setAgtGruppo As enmAgtGruppo 'esecuzione agente gruppo

Public Sub ImpostaAgente(PstrAgente As String, PsetAgtGrp As enmAgtGruppo)
    strAgente = PstrAgente
    setAgtGruppo = PsetAgtGrp
End Sub


Public Function Inizializza(PlngHelpID As Long, _
                            PstrAgente As String, _
                            PsetAgtGruppo As enmAgtGruppo) As Boolean

    Inizializza = False
    If (PlngHelpID <> 0) Then
        Inizializza = True
        lngHelpID = PlngHelpID
        strAgente = PstrAgente
        setAgtGruppo = PsetAgtGruppo
    End If
    
End Function


