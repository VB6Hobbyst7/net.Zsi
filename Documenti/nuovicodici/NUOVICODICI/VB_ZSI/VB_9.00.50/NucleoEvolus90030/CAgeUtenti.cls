VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAgeUtenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'===============================================
'       definizione tipi enumerativi
'===============================================
'Enum enmAgtEventi 'eventi gestiti
'    evtGenerico = 0
'    evtSalvaInserimento = 1
'    evtSalvaModifica = 2
'    evtAnnullamento = 3
'    evtNuovo = 4
'End Enum

'Enum enmAgtGruppo 'gestione agente gruppo
'    agtNonEsegui = 0
'    agtPrima = 1
'    agtDopo = 2
'End Enum

'===============================================
'       definizione variabili
'===============================================
Public vntUserID As Variant
Public setTipoUtente As enmTipoUtente
'Public colImpAgenti As Collection
Public objImpAgenti As CImpostaAge

Private intNrTrm As Integer
Private MstrIniPers As String


Private Sub CaricaImpostazioniGruppo(vntGruppo As Variant, lngIDForm As Long)
    Dim intq As Integer
    Dim strSQL As String
    Dim hSS As CRecordSet
    Dim bolEnd As Boolean
    Dim strAgente As String

    strSQL = "SELECT Agente" _
                & " FROM TabAgentiGruppo " _
                & " WHERE CodGruppo=" & Val(vntGruppo) & " AND FormID=" & lngIDForm
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    If Not bolEnd Then
        strAgente = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Agente", "")
        Call objImpAgenti.ImpostaAgente(strAgente, agtNonEsegui)
        'bolEnd = Not MXDB.dbSuccessivo(hSS, TIPO_SNAPSHOT)
    End If
    intq = MXDB.dbChiudiSS(hSS)

End Sub

Private Sub CaricaImpostazioniUtente(vntUtente As Variant, lngIDForm As Long)
    
    Dim intq As Integer
    Dim strSQL As String
    Dim hSS As CRecordSet
    Dim bolEnd As Boolean
    Dim strAgente As String
    Dim setAgtGrp As enmAgtGruppo

    strSQL = "SELECT Agente, EseguiGruppo" _
                & " FROM TabAgentiUtente" _
                & " WHERE CodUtente='" & vntUtente & "' AND FormID=" & lngIDForm
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    If (Not bolEnd) Then
        strAgente = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Agente", "")
        setAgtGrp = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "EseguiGruppo", agtNonEsegui)
        Call objImpAgenti.ImpostaAgente(strAgente, setAgtGrp)
    End If
    intq = MXDB.dbChiudiSS(hSS)

End Sub


Public Sub CopiaImpostazioni(CSource As CAgeUtenti)

    'cancello gli accessi della destinazione
    With objImpAgenti
        .strAgente = ""
        .setAgtGruppo = agtNonEsegui
    
        'e copio quelli del sorgente
        Call .ImpostaAgente(CSource.objImpAgenti.strAgente, CSource.objImpAgenti.setAgtGruppo)
    End With
End Sub


Public Sub Inizializza(vntUser As Variant, _
                       setTipoUt As enmTipoUtente, _
                       lngIDForm As Long)
    Dim curTag As Integer

    vntUserID = vntUser
    intNrTrm = LeggiNumeroTerminale(vntUser)
    MstrIniPers = MXNU.PercorsoPgm & "\mwpers" & intNrTrm & ".ini"
    setTipoUtente = setTipoUt
    'inizializzo le impostazioni agenti
    
    
    '...e ne leggo le impostazioni
    If (setTipoUt = tipGruppo) Then
        Call CaricaImpostazioniGruppo(vntUser, lngIDForm)
    Else
        Call CaricaImpostazioniUtente(vntUser, lngIDForm)
    End If
    
    
End Sub

Public Sub LeggiImpostazioni(strAgente As String, setAgtGruppo As enmAgtGruppo)
    
    strAgente = objImpAgenti.strAgente
    setAgtGruppo = objImpAgenti.setAgtGruppo
    
End Sub


Public Function SalvaImpostazioniGruppo(lngIDForm As Long) As Boolean
    
    Dim intq As Integer
    Dim strSQL As String
    Dim hDY As CRecordSet
    
    SalvaImpostazioniGruppo = True
    On Local Error GoTo err_SalvaImpostazioniGruppo
    Call MXDB.dbBeginTrans(hndDBArchivi)
    strSQL = " FROM TabAgentiGruppo WHERE CodGruppo=" & Val(vntUserID) & " AND FormID=" & lngIDForm
    'cancello le impostazioni...
    Call MXDB.dbEseguiSQL(hndDBArchivi, "DELETE" & strSQL)
    
    '...e le reinserisco
    Set hDY = MXDB.dbCreaDY(hndDBArchivi, "SELECT CodGruppo,FormID,Agente,UtenteModifica,DataModifica" & strSQL, TIPO_TABELLA)
    
    With objImpAgenti
        If (.strAgente <> "") Then
            Call MXDB.dbInserisci(hDY, TIPO_DYNASET)
            Call MXDB.dbSetCampo(hDY, TIPO_DYNASET, "CodGruppo", Val(vntUserID))
            Call MXDB.dbSetCampo(hDY, TIPO_DYNASET, "FormID", lngIDForm)
            Call MXDB.dbSetCampo(hDY, TIPO_DYNASET, "Agente", .strAgente)
            Call MXDB.dbRegistra(hDY, TIPO_DYNASET)
        End If
    End With
    intq = MXDB.dbChiudiDY(hDY)
    
    Call MXDB.dbCommitTrans(hndDBArchivi)
    
fine_SalvaImpostazioniGruppo:
    On Local Error GoTo 0
Exit Function

err_SalvaImpostazioniGruppo:
    Select Case MXNU.GestioneErrori(Err.Number, "")
        Case NU_ERR_RIPROVA
            Resume
        Case NU_ERR_IGNORA
            Err.Clear
            Resume Next
        Case NU_ERR_RIPORTA
            Call MXDB.dbRollBack(hndDBArchivi)
            On Local Error GoTo 0
            SalvaImpostazioniGruppo = False
            Resume fine_SalvaImpostazioniGruppo
    End Select
    
End Function


Public Function SalvaImpostazioniUtente(lngIDForm As Long) As Boolean
    
    Dim intq As Integer
    Dim strSQL As String
    Dim hDY As CRecordSet
    
    SalvaImpostazioniUtente = True
    On Local Error GoTo err_SalvaImpostazioniUtente
    Call MXDB.dbBeginTrans(hndDBArchivi)
    strSQL = " FROM TabAgentiUtente WHERE CodUtente='" & vntUserID & "' AND FormID=" & lngIDForm
    'cancello le impostazioni...
    Call MXDB.dbEseguiSQL(hndDBArchivi, "DELETE" & strSQL)
    '...e le reinserisco
    Set hDY = MXDB.dbCreaDY(hndDBArchivi, "SELECT CodUtente,FormID,Agente,EseguiGruppo,UtenteModifica,DataModifica" & strSQL, TIPO_TABELLA)
    With objImpAgenti
        If (.strAgente <> "") Then
            Call MXDB.dbInserisci(hDY, TIPO_DYNASET)
            Call MXDB.dbSetCampo(hDY, TIPO_DYNASET, "CodUtente", vntUserID)
            Call MXDB.dbSetCampo(hDY, TIPO_DYNASET, "FormID", lngIDForm)
            Call MXDB.dbSetCampo(hDY, TIPO_DYNASET, "Agente", .strAgente)
            Call MXDB.dbSetCampo(hDY, TIPO_DYNASET, "EseguiGruppo", .setAgtGruppo)
            Call MXDB.dbRegistra(hDY, TIPO_DYNASET)
        End If
    End With
    intq = MXDB.dbChiudiDY(hDY)
    
    Call MXDB.dbCommitTrans(hndDBArchivi)
    
fine_SalvaImpostazioniUtente:
    On Local Error GoTo 0
Exit Function

err_SalvaImpostazioniUtente:
    Select Case MXNU.GestioneErrori(Err.Number, "")
        Case NU_ERR_RIPROVA
            Resume
        Case NU_ERR_IGNORA
            Err.Clear
            Resume Next
        Case NU_ERR_RIPORTA
            Call MXDB.dbRollBack(hndDBArchivi)
            On Local Error GoTo 0
            SalvaImpostazioniUtente = False
            Resume fine_SalvaImpostazioniUtente
    End Select
    
End Function

Public Sub MemorizzaImpostazioni(ByVal strAgente As String, ByVal setAgtGrp As enmAgtGruppo)
    
    Call objImpAgenti.ImpostaAgente(strAgente, setAgtGrp)

End Sub

Private Sub Class_Initialize()
    Set objImpAgenti = New CImpostaAge
End Sub

Private Sub Class_Terminate()
    Set objImpAgenti = Nothing
End Sub


