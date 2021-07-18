VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccUtenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'===============================================
'       definizione tipi enumerativi
'===============================================
Enum enmTipoUtente
    tipUtente = 0
    tipGruppo = 1
End Enum
'===============================================
'       definizione costanti
'===============================================
Private Const OFFSET_SITUAZIONI = 100
'===============================================
'       definizione variabili
'===============================================
Public vntUserID As Variant
Public setTipoUtente As enmTipoUtente
Public colAccessi As Collection

Private intNrTrm As Integer
Private mStrFilePers As String
Private mStrNomeForm As String

'==============================================================================================
'               funzioni private della classe
'==============================================================================================
Private Sub InizializzaAccessi(lngIDForm As Long, vetIDSchede() As Integer)
Dim cnt As Integer
Dim CAcc As CAccessi

    'accessi form
    Set CAcc = New CAccessi
    Call CAcc.Inizializza(lngIDForm, ID_SCHEDA_FORM, ACC_NONDEFINITO)
    Call colAccessi.Add(CAcc, KeyAccessoGet(lngIDForm, ID_SCHEDA_FORM))
    'accessi schede
    For cnt = 0 To UBound(vetIDSchede)
        Set CAcc = New CAccessi
        Call CAcc.Inizializza(lngIDForm, vetIDSchede(cnt), ACC_NONDEFINITO)
        Call colAccessi.Add(CAcc, KeyAccessoGet(lngIDForm, vetIDSchede(cnt)))
    Next cnt
End Sub

Private Sub SitInizializzaAccessi(vetIDVisioni() As Integer)
Dim cnt As Integer
Dim CAcc As CAccessi
    
    'accessi situazioni
    For cnt = 0 To UBound(vetIDVisioni)
        Set CAcc = New CAccessi
        Call CAcc.Inizializza(ID_SCHEDA_SITUAZIONE, vetIDVisioni(cnt), ACC_NONDEFINITO)
        Call colAccessi.Add(CAcc, KeyAccessoGet(ID_SCHEDA_SITUAZIONE, vetIDVisioni(cnt)))
    Next cnt
End Sub

Private Sub CaricaAccessi(hSS As MXKit.CRecordSet, lngIDForm As Long)
Dim bolEnd As Boolean
Dim intScheda As Integer
Dim intAccesso As Integer
Dim strKey As String

    bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    Do While (Not bolEnd)
        intScheda = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "IndiceScheda", 0)
        intAccesso = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "TipoAccesso", ACC_NONDEFINITO)
        If intScheda <= -OFFSET_SITUAZIONI Then
            intScheda = intScheda + OFFSET_SITUAZIONI
            'RIF.A#9769 - controllo che l'elemento della collezione esista
            strKey = KeyAccessoGet(ID_SCHEDA_SITUAZIONE, intScheda)
            If (EsisteElementoCollection(colAccessi, strKey)) Then
                Call colAccessi(strKey).ImpostaAccessi(intAccesso)
            End If
        Else
            'RIF.A#9769 - controllo che l'elemento della collezione esista
            strKey = KeyAccessoGet(lngIDForm, intScheda)
            If (EsisteElementoCollection(colAccessi, strKey)) Then
                Call colAccessi(strKey).ImpostaAccessi(intAccesso)
            End If
        End If
        bolEnd = Not MXDB.dbSuccessivo(hSS, TIPO_SNAPSHOT)
    Loop
End Sub

Private Sub CaricaAccessiGruppo(vntGruppo As Variant, lngIDForm As Long)
Dim strSQL As String
Dim hSS As CRecordSet

    strSQL = "SELECT IndiceScheda,TipoAccesso" _
                & " FROM TabAccessiGruppo " _
                & " WHERE CodGruppo=" & Val(vntGruppo) & " AND HelpID=" & lngIDForm
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    Call CaricaAccessi(hSS, lngIDForm)
    Call MXDB.dbChiudiSS(hSS)
End Sub

Private Sub CaricaAccessiUtente(vntUtente As Variant, lngIDForm As Long)
Dim strSQL As String
Dim hSS As CRecordSet

    strSQL = "SELECT IndiceScheda,TipoAccesso" _
                & " FROM TabAccessiUtente " _
                & " WHERE CodUtente='" & vntUtente & "' AND HelpID=" & lngIDForm
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    Call CaricaAccessi(hSS, lngIDForm)
    Call MXDB.dbChiudiSS(hSS)
End Sub

Private Function KeyAccessoGet(lngHelpID As Long, intScheda As Integer) As Variant
    KeyAccessoGet = CStr(lngHelpID) & "." & CStr(intScheda)
End Function

'Private Sub SitCaricaAccessi(strSitEntry As String, vetIDVisioni() As Integer)
'Dim strAcc As String
'Dim strRic As String
'Dim intSit As Integer
'
'    'carica accessi situazioni
'    strAcc = MXNU.LeggiProfilo(mStrFilePers, "SITUAZIONI", strSitEntry, "")
''    If Len(strAcc) > 0 Then
'        strAcc = ";" & strAcc & ";"
'        For intSit = 1 To UBound(vetIDVisioni)
'            strRic = ";" & CStr(Abs(vetIDVisioni(intSit))) & ";"
'            If (InStr(strAcc, strRic) = 0) Then
'                Call colAccessi(KeyAccessoGet(ID_SCHEDA_SITUAZIONE, vetIDVisioni(intSit))).ImpostaAccessi(ACC_NESSUNO)
'            End If
'        Next intSit
''    End If
'End Sub

Private Sub Class_Initialize()
    Set colAccessi = New Collection
End Sub

Private Sub Class_Terminate()
    Set colAccessi = Nothing
End Sub

'==============================================================================================
'               funzioni pubbliche della classe
'==============================================================================================
Public Function AccessiImpostati(ByVal lngIDForm As Long, ByVal intIDScheda As Integer) As Boolean
    On Local Error Resume Next
    With colAccessi(KeyAccessoGet(lngIDForm, intIDScheda))
        If (Err = 0) Then
            AccessiImpostati = (.intAccesso <> ACC_NONDEFINITO)
        Else
            AccessiImpostati = False
        End If
    End With
    On Local Error GoTo 0
End Function

Public Sub CopiaAccessi(CSource As CAccUtenti)
Dim CAcc As CAccessi

    'cancello gli accessi della destinazione
    For Each CAcc In colAccessi
        CAcc.intAccesso = ACC_NONDEFINITO
    Next
    'e copio quelli del sorgente
    For Each CAcc In colAccessi
        Call CAcc.ImpostaAccessi(CSource.colAccessi(KeyAccessoGet(CAcc.lngHelpID, CAcc.intScheda)).intAccesso)
    Next CAcc

End Sub

Public Sub Inizializza(vntUser As Variant, _
                        setTipoUt As enmTipoUtente, _
                        lngIDForm As Long, _
                        vetIDSchede() As Integer, _
                        bolSituazione As Boolean, _
                        vetIDVisioni() As Integer, _
                        strSitEntry As String, _
                        strNomeForm As String)

    'imposto utente
    vntUserID = vntUser
    intNrTrm = LeggiNumeroTerminale(vntUser)
    mStrFilePers = MXNU.PercorsoPreferenze & "\mwpers" & intNrTrm & ".ini"
    setTipoUtente = setTipoUt
    mStrNomeForm = strNomeForm
    'inizializzo gli accessi...
    Call InizializzaAccessi(lngIDForm, vetIDSchede())
    If (bolSituazione) Then Call SitInizializzaAccessi(vetIDVisioni())
    '...e ne leggo le impostazioni
    If (setTipoUtente = tipGruppo) Then
        Call CaricaAccessiGruppo(vntUser, lngIDForm)
    Else
        Call CaricaAccessiUtente(vntUser, lngIDForm)
    End If
End Sub

Public Sub LeggiAccessi(ByVal lngIDForm As Long, _
                            ByVal intIDScheda As Integer, _
                            bolAccesso As Boolean, _
                            Optional bolLettura As Boolean, _
                            Optional bolModifica As Boolean, _
                            Optional bolInserisci As Boolean, _
                            Optional bolAnnulla As Boolean)

Dim intAccessi As Integer

    On Local Error Resume Next
    'leggo il valore impostato
    intAccessi = colAccessi(KeyAccessoGet(lngIDForm, intIDScheda)).intAccesso
    'imposto gli accessi
    bolAccesso = (intAccessi <> ACC_NESSUNO)
    bolLettura = (intAccessi And ACC_LETTURA) = ACC_LETTURA
    bolModifica = (intAccessi And ACC_MODIFICA) = ACC_MODIFICA
    bolInserisci = (intAccessi And ACC_INSERISCI) = ACC_INSERISCI
    bolAnnulla = (intAccessi And ACC_ANNULLA) = ACC_ANNULLA
    On Local Error GoTo 0
End Sub

Public Sub MemorizzaAccessi(ByVal lngIDForm As Long, _
                            ByVal intIDScheda As Integer, _
                            ByVal bolImpostato As Boolean, _
                            ByVal bolAccesso As Boolean, _
                            ByVal bolLettura As Boolean, _
                            ByVal bolModifica As Boolean, _
                            ByVal bolInserisci As Boolean, _
                            ByVal bolAnnulla As Boolean)

Dim intAccessi As Integer

    'calcolo il valore degli accessi
    If (Not bolImpostato) Then
        intAccessi = ACC_NONDEFINITO
    Else
        intAccessi = ACC_NESSUNO
        If (bolAccesso) Then
            If bolLettura Then intAccessi = intAccessi + ACC_LETTURA
            If bolModifica Then intAccessi = intAccessi + ACC_MODIFICA
            If bolInserisci Then intAccessi = intAccessi + ACC_INSERISCI
            If bolAnnulla Then intAccessi = intAccessi + ACC_ANNULLA
        End If
    End If
    'memorizzo i valori impostati
    colAccessi(KeyAccessoGet(lngIDForm, intIDScheda)).intAccesso = intAccessi

End Sub

Private Sub SalvaAccessi(hDY As MXKit.CRecordSet, ByVal lngIDForm As Long)
    On Local Error GoTo SalvaAccessi_ERR

Dim CAcc As CAccessi
Dim intScheda As Integer
Dim strDes As String

    With MXDB
        For Each CAcc In colAccessi
            If (CAcc.intAccesso <> ACC_NONDEFINITO) Then 'inserisco solo record con limitazioni
                Call .dbInserisci(hDY, TIPO_DYNASET)
                Call .dbSetCampo(hDY, TIPO_DYNASET, 0, vntUserID) 'codutente/codgruppo
                Call .dbSetCampo(hDY, TIPO_DYNASET, "HelpID", lngIDForm)
                intScheda = CAcc.intScheda
                If (CAcc.lngHelpID = ID_SCHEDA_SITUAZIONE) Then
                    intScheda = intScheda - OFFSET_SITUAZIONI
                End If
                Call .dbSetCampo(hDY, TIPO_DYNASET, "IndiceScheda", intScheda)
                Call .dbSetCampo(hDY, TIPO_DYNASET, "TipoAccesso", CAcc.intAccesso)
                Call .dbRegistra(hDY, TIPO_DYNASET)
                
                If Not MXNU.CtrlAccessi Then
                    If CAcc.intAccesso = 0 Then
                        Call MXNU.WritePrivacyLog(DisabilitaUtente, MXNU.CaricaStringaRes(3208, Array("", vntUserID, mStrNomeForm, intScheda)))
                    Else
                        strDes = ""
                        If (CAcc.intAccesso And ACC_LETTURA) = ACC_LETTURA Then strDes = MXNU.CaricaStringaRes(3213)
                        If (CAcc.intAccesso And ACC_MODIFICA) = ACC_MODIFICA Then strDes = strDes & " " & MXNU.CaricaStringaRes(3214)
                        If (CAcc.intAccesso And ACC_INSERISCI) = ACC_INSERISCI Then strDes = strDes & " " & MXNU.CaricaStringaRes(3215)
                        If (CAcc.intAccesso And ACC_ANNULLA) = ACC_ANNULLA Then strDes = strDes & " " & MXNU.CaricaStringaRes(3216)
                        Call MXNU.WritePrivacyLog(AbilitaUtente, MXNU.CaricaStringaRes(3207, Array("", vntUserID, mStrNomeForm, intScheda, strDes)))
                    End If
                End If
            End If
        Next
    End With
    
SalvaAccessi_END:
    On Local Error GoTo 0
    Exit Sub
    
SalvaAccessi_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Err.Raise lngErrCod, strErrDsc
End Sub

Public Function SalvaAccessiGruppo(lngIDForm As Long) As Boolean
Dim strSQL As String
Dim hDY As CRecordSet

    SalvaAccessiGruppo = True
    On Local Error GoTo err_SalvaAccessiGruppo
    Call MXDB.dbBeginTrans(hndDBArchivi)
    strSQL = " FROM TabAccessiGruppo WHERE CodGruppo=" & Val(vntUserID) & " AND HelpID=" & lngIDForm
    'cancello gli accessi...
    Call MXDB.dbEseguiSQL(hndDBArchivi, "DELETE" & strSQL)
    '...e li reinserisco
    Set hDY = MXDB.dbCreaDY(hndDBArchivi, "SELECT CodGruppo,HelpID,IndiceScheda,TipoAccesso,UtenteModifica,DataModifica" & strSQL, TIPO_TABELLA)
    Call SalvaAccessi(hDY, lngIDForm)
    Call MXDB.dbChiudiDY(hDY)
    Call MXDB.dbCommitTrans(hndDBArchivi)
    
fine_SalvaAccessiGruppo:
    On Local Error GoTo 0
Exit Function

err_SalvaAccessiGruppo:
    On Local Error GoTo 0
    Call MXDB.dbRollBack(hndDBArchivi)
    SalvaAccessiGruppo = False
    Resume fine_SalvaAccessiGruppo
End Function

Public Function SalvaAccessiUtente(lngIDForm As Long, strSitEntry As String) As Boolean
Dim strSQL As String
Dim hDY As CRecordSet

    SalvaAccessiUtente = True
    On Local Error GoTo err_SalvaAccessiUtente
    Call MXDB.dbBeginTrans(hndDBArchivi)
    strSQL = " FROM TabAccessiUtente WHERE CodUtente='" & vntUserID & "' AND HelpID=" & lngIDForm
    'cancello gli accessi...
    Call MXDB.dbEseguiSQL(hndDBArchivi, "DELETE" & strSQL)
    '...e li reinserisco
    Set hDY = MXDB.dbCreaDY(hndDBArchivi, "SELECT CodUtente,HelpID,IndiceScheda,TipoAccesso,UtenteModifica,DataModifica" & strSQL, TIPO_TABELLA)
    Call SalvaAccessi(hDY, lngIDForm)
    Call MXDB.dbChiudiDY(hDY)
    Call MXDB.dbCommitTrans(hndDBArchivi)
        
fine_SalvaAccessiUtente:
    On Local Error GoTo 0
Exit Function

err_SalvaAccessiUtente:
    On Local Error GoTo 0
    Call MXDB.dbRollBack(hndDBArchivi)
    SalvaAccessiUtente = False
    Resume fine_SalvaAccessiUtente
End Function
