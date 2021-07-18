VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVisioniConSelez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum setTipoVisioneConSelez
    VCS_DICHINTENTO = 0
    VCS_TRASFDEFINITIVE = 1
    VCS_MANUTEFFETTI = 2
    VCS_CANCSCAD = 3
    VCS_RICDOC = 4
    VCS_CANCDOC = 5
    VCS_PIANIFICAZIONE = 6
    VCS_MPSACCODAPERIODI = 7
    VCS_MPSCHIUDIPERIODI = 8
    VCS_MPSGENERAPIANO = 9
    VCS_ANACONS = 10
    VCS_STORALLEST = 11
    VCS_AGGMAGVBANCO = 12
    VCS_CONTCORRVBANCO = 13
    VCS_GENDOCRIEPVBANCO = 14
    VCS_COSTIPRODOTTO = 15
    VCS_GESTPROD = 16
    VCS_ANNULLACICLI = 17
    VCS_RICMOVCONSUNT = 18
    VCS_RIVCICLIDACOMM = 19
    VCS_CONSUNTCOMMCLI = 20
    VCS_AGGCAMPIARTFIGLI = 21
    VCS_GENERACONSUMICLAV = 22
End Enum

Private mvarTipoElaborazione As setTipoVisioneConSelez
Private mvarElabInterrotta As Boolean

Public Function CambiaVisione(intListIndex As Integer) As String

End Function

Public Property Let TipoElaborazione(ByVal vData As setTipoVisioneConSelez)
    mvarTipoElaborazione = vData
End Property

Public Property Let ElabInterrotta(ByVal vData As Boolean)
    mvarElabInterrotta = vData
End Property

Public Property Get TipoElaborazione() As setTipoVisioneConSelez
    TipoElaborazione = mvarTipoElaborazione
End Property

Public Property Get ElabInterrotta() As Boolean
    ElabInterrotta = mvarElabInterrotta
End Property

Public Sub UsaCollection(strDati As Collection, intListIndex As Integer, ssFiltroDati As Object, ssSpread As Object)
    Dim objUsaCollection As Object
End Sub

Public Sub InitTabella(xTabella As MXKit.CTabelle, intSelTabella As Integer, ssFoglio As FPSpreadADO.fpSpread)

End Sub
