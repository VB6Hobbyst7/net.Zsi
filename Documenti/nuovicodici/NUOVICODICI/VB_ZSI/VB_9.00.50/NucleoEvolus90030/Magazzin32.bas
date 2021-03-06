Attribute VB_Name = "MMagazzino"
Option Explicit
DefLng A-Z

'================================
'   definizione costanti
'================================
Const MAG_DFLT_LenTiplogia = 3 'lunghezza tipologia
Const MAG_DFLT_LenVariante = 8 'lunghezza variante
Const MAG_DFLT_LenArticolo = 50 'lunghezza articolo
Const MAG_DFLT_LenPartita = 15 'lunghezza partita
Const MAG_DFLT_LenUM = 3 'lunghezza unità di misura
Const MAG_DFLT_DecFC = 9 'numero decimali fattore conversione
Global Const MAG_TUTTE_LE_PARTITE = "ĦĦĦĦĦĦĦĦĦĦĦĦĦĦĦ"
Global Const MAG_TUTTE_LE_UBICAZIONI = "ĦĦĦĦĦĦĦĦĦĦĦĦĦĦĦ"
'================================
'   definizione tipi enumerativi
'================================
'provenienza dell'articolo
Public Enum setProvenienzaArticolo
    PA_daAcquisto = 0
    PA_daProduzione = 1
    PA_daContoLavoro = 2
End Enum
'arrotondamento lead time
Public Enum setArrotondaLeadTime
    arrLTProporzionale = 0
    arrLTMultiplo = 1
End Enum
'dati articolo
Public Enum setDatiArticolo
    artNessuna = 0
    artAnagrafici = 1
    artCommerciali = 2
    artProduzione = 3
    artInformazioni = 4
    artLIFO = 5
    artExtra = 6
End Enum

'flag sui vincoli per la generazione della partita
Public Enum setFlagGeneraPartita
    partitaRichiedi = 0
    partitaNonInserire = 1
    partitaInserisciDefault = 2
End Enum
'================================
'   definizione variabili
'================================

'NOME           : LeggiVincoliMagazzino
'DESCRIZIONE    : legge i vincoli e le dimensioni dei campi del magazzino
Sub LeggiVincoliMagazzino()
Dim hSS As CRecordSet
Dim strSQL As String
    With MXDB
        'leggo le dimensioni del campo tipologia
        strSQL = "SELECT Tipologia FROM TabTipologie WHERE Tipologia=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        MXNU.MAG_LenTiplogia = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenTiplogia = 0) Then MXNU.MAG_LenTiplogia = MAG_DFLT_LenTiplogia
        Call .dbChiudiSS(hSS)
        'leggo le dimensioni del campo variante
        strSQL = "SELECT Variante FROM TabVarianti WHERE Variante=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        MXNU.MAG_LenVariante = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenVariante = 0) Then MXNU.MAG_LenVariante = MAG_DFLT_LenVariante
        Call .dbChiudiSS(hSS)
        'leggo le dimensioni del campo codice articolo
        strSQL = "SELECT Codice, Descrizione FROM AnagraficaArticoli WHERE Codice=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        MXNU.MAG_LenArticolo = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenArticolo = 0) Then MXNU.MAG_LenArticolo = MAG_DFLT_LenArticolo
        'RIF.A#11745 - leggo dimensione massima descrizione articolo
        MXNU.MAG_LenDscArticolo = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 1)
        If (MXNU.MAG_LenDscArticolo = 0) Then MXNU.MAG_LenDscArticolo = 80
        
        Call .dbChiudiSS(hSS)
        'leggo le dimensioni del campo partita
        strSQL = "SELECT CodLotto FROM AnagraficaLotti WHERE CodArticolo=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        MXNU.MAG_LenPartita = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenPartita = 0) Then MXNU.MAG_LenPartita = MAG_DFLT_LenPartita
        Call .dbChiudiSS(hSS)
        'leggo le dimensioni del campo unità di misura
        strSQL = "SELECT Codice FROM TabUnitaMisura WHERE Codice=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        MXNU.MAG_LenUM = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenUM = 0) Then MXNU.MAG_LenUM = MAG_DFLT_LenUM
        Call .dbChiudiSS(hSS)
        'leggo il numero decimali del campo fattore conversione
        strSQL = "SELECT Fattore FROM ArticoliFattoriConversione WHERE CodArt=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strSQL)
        Call .dbGetLenCampo(hSS, TIPO_SNAPSHOT, "Fattore", MXNU.MAG_DecFC)
        'ATTENZIONE: mettendo 10 come decimali fattore conversione fdec tronca ad una cifra decimale
        If (MXNU.MAG_DecFC = 0 Or MXNU.MAG_DecFC > 9) Then MXNU.MAG_DecFC = MAG_DFLT_DecFC
        Call .dbChiudiSS(hSS)
    End With
End Sub


'NOME           : ArticoloMovimentato
'DESCRIZIONE    : controlla se ci sono movimenti di storico che fanno riferimento all'articolo
'PARAMETRO 1    : articolo da controllare
Function ArticoloMovimentato(ByVal strCodArt As String) As Boolean
Dim strSQL As String
Dim hSS As MXKit.CRecordSet

    With MXDB
        strSQL = "SELECT TOP 1 Progressivo" _
            & " FROM StoricoMag" _
            & " WHERE CodArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
        Set hSS = .dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        ArticoloMovimentato = (Not .dbFineTab(hSS, TIPO_SNAPSHOT))
        Call .dbChiudiSS(hSS)
    End With

End Function

'NOME           : ArticoloGeneratore
'DESCRIZIONE    : controlla se per un articolo con tipologie ci sono articoli a varianti generati
'PARAMETRO 1    : articolo con tipologie da controllare
Function ArticoloGeneratore(ByVal strCodArt As String) As Boolean
Dim strSQL As String
Dim hSS As MXKit.CRecordSet

    With MXDB
        strSQL = "SELECT Codice" _
            & " FROM AnagraficaArticoli" _
            & " WHERE CodicePrimario=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
        Set hSS = .dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        ArticoloGeneratore = (Not .dbFineTab(hSS, TIPO_SNAPSHOT))
        Call .dbChiudiSS(hSS)
    End With

End Function

Function CaricaDatiArticolo(ByVal vntArticolo As Variant, _
                            ByVal strListaCampi As String, _
                            ByVal enmTipoDato As setDatiArticolo, _
                            colValoriRitorno As Collection, _
                            Optional bolLettiDatiPadre As Boolean = False) As Boolean
                            
    If (MXDB.SupportEnhancements) Then
        CaricaDatiArticolo = CaricaDatiArticoloExt(vntArticolo, strListaCampi, enmTipoDato, colValoriRitorno, bolLettiDatiPadre)
    Else
        CaricaDatiArticolo = CaricaDatiArticoloOld(vntArticolo, strListaCampi, enmTipoDato, colValoriRitorno, bolLettiDatiPadre)
    End If
                            
End Function

'NOME           : CaricaDatiArticolo
'DESCRIZIONE    : carica i dati di un articolo e, se non generato, dall'articolo con tipologia
'PARAMETRO 1    : codice articolo
'PARAMETRO 2    : lista campi da leggere
'PARAMETRO 3    : collection campi ritorno
'PARAMETRO 4    : flag carica i dati dell'articolo padre (nel caso di art. varianti non generato) si/no
'RISULTATO      : esito del caricamento
'ATTENZIONE     : la funzione non carica i campi CODICE e DESCRIZIONE
Private Function CaricaDatiArticoloOld(ByVal vntArticolo As Variant, _
                            ByVal strListaCampi As String, _
                            ByVal enmTipoDato As setDatiArticolo, _
                            colValoriRitorno As Collection, _
                            Optional bolLettiDatiPadre As Boolean = False) As Boolean
                            
Dim strSQL As String
Dim strFrom As String, strWhr As String
Dim hSS As CRecordSet
Dim intPosSep As Integer
Dim cnt As Integer, intNC As Integer
Dim vetCampi() As String
Dim strCodArt As String
    
    CaricaDatiArticoloOld = True
    bolLettiDatiPadre = False
    'leggo i campi aggiuntivi
    strCodArt = CStr(vntArticolo)
    If (Len(strCodArt) = 0 Or Len(strListaCampi) = 0) Then
        'rif.A-3707 - NECESSARIO in quanto se non passo il codice o la lista campi utilizzo la collection
        '             dei valori di ritorno che risulterà senza elementi e pertanto l'accesso va in errore
        CaricaDatiArticoloOld = False
    Else
        GoSub componiQuery
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        If MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
            bolLettiDatiPadre = True
            'cerco su codice padre
            intPosSep = InStr(vntArticolo, MXNU.SepVar)
            If (intPosSep <> 0) Then
                Call MXDB.dbChiudiSS(hSS)
                strCodArt = Left$(vntArticolo, intPosSep - 1)
                GoSub componiQuery
                Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
            End If
        End If
        If (MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
            CaricaDatiArticoloOld = False
        Else
            ReDim vetCampi(0) As String
            intNC = slice(strListaCampi, ",", vetCampi())
            For cnt = 0 To intNC - 1
                If (StrComp(vetCampi(cnt), "codice", vbTextCompare) <> 0) Then
                    On Local Error Resume Next
                    Call colValoriRitorno.Add(MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, vetCampi(cnt), Empty), vetCampi(cnt))
                    On Local Error GoTo 0
                End If
            Next cnt
        End If
    End If
    
fine_CaricaDatiArticolo:
    Call MXDB.dbChiudiSS(hSS)
    Exit Function
    
componiQuery:
    Select Case enmTipoDato
        Case artNessuna
            'OTTIMIZZAZIONE: risulta migliore che utilizzare VISTAANAGRAFICAARTICOLI
            strFrom = "{oj ANAGRAFICAARTICOLI ART inner join" _
                    & " ANAGRAFICAARTICOLICOMM COMM on ART.CODICE = COMM.CODICEART inner join" _
                    & " ANAGRAFICAARTICOLIPROD PROD on ART.CODICE = PROD.CODICEART}"
            strWhr = "ART.CODICE=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT) _
                    & " and COMM.ESERCIZIO=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER) _
                    & " and PROD.ESERCIZIO=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER)
        Case artAnagrafici
            strFrom = "AnagraficaArticoli"
            strWhr = "Codice=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
        Case artCommerciali
            strFrom = "AnagraficaArticoliComm"
            strWhr = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER)
        Case artInformazioni
            strFrom = "DescrArticoli"
            strWhr = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
        Case artProduzione
            strFrom = "AnagraficaArticoliProd"
            strWhr = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER)
        Case artLIFO
            strFrom = "LifoArticoli"
            strWhr = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER)
        Case artExtra
            strFrom = "ExtraMag"
            strWhr = "CodArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
    End Select
    strSQL = "SELECT " & strListaCampi _
            & " FROM " & strFrom _
            & " WHERE " & strWhr
Return
End Function

'NOME           : CaricaDatiArticolo
'DESCRIZIONE    : carica i dati di un articolo e, se non generato, dall'articolo con tipologia
'PARAMETRO 1    : codice articolo
'PARAMETRO 2    : lista campi da leggere
'PARAMETRO 3    : collection campi ritorno
'PARAMETRO 4    : flag carica i dati dell'articolo padre (nel caso di art. varianti non generato) si/no
'RISULTATO      : esito del caricamento
'ATTENZIONE     : la funzione non carica i campi CODICE e DESCRIZIONE
Private Function CaricaDatiArticoloExt(ByVal vntArticolo As Variant, _
                            ByVal strListaCampi As String, _
                            ByVal enmTipoDato As setDatiArticolo, _
                            colValoriRitorno As Collection, _
                            Optional bolLettiDatiPadre As Boolean = False) As Boolean
    
Dim strSQL As MXKit.StatementFragment
Dim strFrom As MXKit.StatementFragment, strWhr As MXKit.StatementFragment
Dim hSS As CRecordSet
Dim intPosSep As Integer
Dim cnt As Integer, intNC As Integer
Dim vetCampi() As String
Dim strCodArt As String
    
    CaricaDatiArticoloExt = True
    bolLettiDatiPadre = False
    'leggo i campi aggiuntivi
    strCodArt = CStr(vntArticolo)
    If (Len(strCodArt) = 0 Or Len(strListaCampi) = 0) Then
        'rif.A-3707 - NECESSARIO in quanto se non passo il codice o la lista campi utilizzo la collection
        '             dei valori di ritorno che risulterà senza elementi e pertanto l'accesso va in errore
        CaricaDatiArticoloExt = False
    Else
        GoSub componiQuery
        Set hSS = MXDB.dbCreaSSEx(hndDBArchivi, strSQL, TIPO_TABELLA)
        If MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
            bolLettiDatiPadre = True
            'cerco su codice padre
            intPosSep = InStr(vntArticolo, MXNU.SepVar)
            If (intPosSep <> 0) Then
                Call MXDB.dbChiudiSS(hSS)
                strCodArt = Left$(vntArticolo, intPosSep - 1)
                GoSub componiQuery
                Set hSS = MXDB.dbCreaSSEx(hndDBArchivi, strSQL, TIPO_TABELLA)
            End If
        End If
        If (MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
            CaricaDatiArticoloExt = False
        Else
            ReDim vetCampi(0) As String
            intNC = slice(strListaCampi, ",", vetCampi())
            For cnt = 0 To intNC - 1
                If (StrComp(vetCampi(cnt), "codice", vbTextCompare) <> 0) Then
                    On Local Error Resume Next
                    Call colValoriRitorno.Add(MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, vetCampi(cnt), Empty), vetCampi(cnt))
                    On Local Error GoTo 0
                End If
            Next cnt
        End If
    End If
    
fine_CaricaDatiArticolo:
    Set strSQL = Nothing
    Call MXDB.dbChiudiSS(hSS)
    Exit Function
    
componiQuery:
    Set strFrom = New StatementFragment
    Set strWhr = New StatementFragment
    Select Case enmTipoDato
        Case artNessuna
            'OTTIMIZZAZIONE: risulta migliore che utilizzare VISTAANAGRAFICAARTICOLI
            strFrom.Statement = "{oj ANAGRAFICAARTICOLI ART inner join" _
                    & " ANAGRAFICAARTICOLICOMM COMM on ART.CODICE = COMM.CODICEART inner join" _
                    & " ANAGRAFICAARTICOLIPROD PROD on ART.CODICE = PROD.CODICEART}"
            strWhr.Statement = "ART.CODICE=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo)) _
                    & " and COMM.ESERCIZIO=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO1", adDecimal, 5, adParamInput, 5, 0) _
                    & " and PROD.ESERCIZIO=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO2", adDecimal, 5, adParamInput, 5, 0)
        Case artAnagrafici
            strFrom.Statement = "AnagraficaArticoli"
            strWhr.Statement = "Codice=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo))
        Case artCommerciali
            strFrom.Statement = "AnagraficaArticoliComm"
            strWhr.Statement = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo)) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO1", adDecimal, 5, adParamInput, 5, 0)
        Case artInformazioni
            strFrom.Statement = "DescrArticoli"
            strWhr.Statement = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo))
        Case artProduzione
            strFrom.Statement = "AnagraficaArticoliProd"
            strWhr.Statement = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo)) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO1", adDecimal, 5, adParamInput, 5, 0)
        Case artLIFO
            strFrom.Statement = "LifoArticoli"
            strWhr.Statement = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo)) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO1", adDecimal, 5, adParamInput, 5, 0)
        Case artExtra
            strFrom.Statement = "ExtraMag"
            strWhr.Statement = "CodArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo))
    End Select
    Set strSQL = Nothing
    Set strSQL = New MXKit.StatementFragment
    strSQL.AppendFragments "SELECT " & strListaCampi & " FROM ", strFrom, " WHERE ", strWhr
    
    Set strFrom = Nothing
    Set strWhr = Nothing
Return
End Function

Function GeneraArticoloVarianti(ByVal strCodArt As String, _
    Optional strDscArt As String = "", _
    Optional strVarEspl As String = "", _
    Optional bolAggMag As Boolean = False, _
    Optional bolCopiaExtra As Boolean = True) As Boolean
    
Dim intq As Integer
Dim xCArt As MXBusiness.CVArt
Dim strSQL As String
Dim hSS As MXKit.CRecordSet

    GeneraArticoloVarianti = True
    bolAggMag = False
    'inizializzo le classi
    strSQL = "SELECT AggiornaMag" _
        & " FROM AnagraficaArticoli" _
        & " WHERE Codice=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    If (MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
        Call MXDB.dbChiudiSS(hSS)
        'il codice non esiste -> lo genero
        Set xCArt = MXART.CreaCVArt()
        With xCArt
            .Codice = strCodArt
            If (Not .Valida(CHIEDIVAR_NESSUNA, False, , 0, False)) Then
                GeneraArticoloVarianti = False
                GoTo fine_GeneraArticoloVarianti
            Else
                GeneraArticoloVarianti = .Genera(bolCopiaExtra)
                'rileggo il flag aggiorna magazzino
                Call MXDB.dbChiudiSS(hSS)
                strSQL = "SELECT AggiornaMag" _
                    & " FROM AnagraficaArticoli" _
                    & " WHERE Codice=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
                Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
            End If
        End With
    End If
    'restituisco il flag aggiorna magazzino
    bolAggMag = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "AggiornaMag", True)
    
fine_GeneraArticoloVarianti:
    On Local Error GoTo 0
    GoSub disalloca_GeneraArticoloVarianti
Exit Function

disalloca_GeneraArticoloVarianti:
    'disalloco variabili
    If Not (xCArt Is Nothing) Then
        Call xCArt.Termina
    End If
    Set xCArt = Nothing
    If Not (hSS Is Nothing) Then Call MXDB.dbChiudiDY(hSS)
Return

err_GeneraArticoloVarianti:
    GeneraArticoloVarianti = False
    Call MXNU.MsgBoxEX(1866, vbCritical, 1007, Array(Err.Number, Err.Description, strCodArt))
    Resume fine_GeneraArticoloVarianti:

End Function

'************************************************************************
'NOME           : ArticoloCancellabile
'DESCRIZIONE    : controlla se un articolo è o meno cancellabile
'PARAMETRO 1    : codice articolo
'PARAMETRO 2    : flag articolo tipologia
'************************************************************************
Function ArticoloCancellabile(strCodArt As String, bolArtTipologia As Boolean) As Boolean
Dim strSQL As String
Dim hSS As CRecordSet
Dim strMsg As String

    ArticoloCancellabile = True
    strMsg = ""
    If bolArtTipologia Then
        'controllo se ci sono articoli generati
        strSQL = "SELECT Codice" _
                & " FROM AnagraficaArticoli" _
                & " WHERE (CodicePrimario='" & strCodArt & "')"
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        If (Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
            strMsg = MXNU.CaricaStringaRes(1852, Array("", strCodArt))
            GoTo err_ArticoloCancellabile
        End If
        Call MXDB.dbChiudiSS(hSS)
    Else
        'controllo movimenti di magazzino
        strSQL = "SELECT TOP 1 Progressivo" _
                & " FROM StoricoMag" _
                & " WHERE CodArt='" & strCodArt & "'"
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        If (Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
            strMsg = MXNU.CaricaStringaRes(1853, Array("", strCodArt))
            Call MXDB.dbChiudiSS(hSS)
            GoTo err_ArticoloCancellabile
        End If
        Call MXDB.dbChiudiSS(hSS)
    End If
    'controllo esistenza distinta
'    If (InStr(strCodArt, "#") > 0 Or bolArtTipologia) Then
'        strSQL = "SELECT Progressivo" _
'                & " FROM DistintaArtComposti" _
'                & " WHERE (ArtComposto = '" & strCodArt & "')"
'    Else
'        strSQL = "SELECT Progressivo" _
'                & " FROM DistintaArtComposti" _
'                & " WHERE (ArtComposto = '" & strCodArt & "') OR (ArtComposto = '" & strCodArt & "#')"
'    End If
'    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
'    If (Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
'        strMsg = MXNU.CaricaStringaRes(1854, Array("", strCodArt))
'        GoTo err_ArticoloCancellabile
'    End If
'    Call MXDB.dbChiudiSS(hSS)
        
fine_ArticoloCancellabile:
    Exit Function
    
err_ArticoloCancellabile:
    ArticoloCancellabile = False
    Call MXNU.MsgBoxEX(strMsg, vbExclamation, 1007)
    GoTo fine_ArticoloCancellabile
    
End Function

'NOME           : LeggiMagPrincipale
'DESCRIZIONE    : legge il magazzino principale
'PARAMETRO 1    : (ritorno) codice magazzino principale
'PARAMETRO 2    : (ritorno) descrizione magazzino principale
'RITORNO        : True se il magazzino principale esiste, False altrimenti
Function LeggiMagPrincipale(strCodMagP As String, strDscMagP As String) As Boolean
Dim strSQL As String
Dim hSS As CRecordSet

    strSQL = "SELECT Codice,Descrizione" _
            & " FROM AnagraficaDepositi" _
            & " WHERE Principale <> 0"
    Set hSS = MXDB.dbCreaDY(hndDBArchivi, strSQL, TIPO_TABELLA)
    LeggiMagPrincipale = Not MXDB.dbFineTab(hSS, TIPO_DYNASET)
    strCodMagP = MXDB.dbGetCampo(hSS, TIPO_DYNASET, "Codice", "")
    strDscMagP = MXDB.dbGetCampo(hSS, TIPO_DYNASET, "Descrizione", "")
    Call MXDB.dbChiudiDY(hSS)
End Function

'NOME           : CaricaComboRaggruppaProd
'DESCRIZIONE    : carica il combo dei raggruppamento di produzione e/o assegna il valore a tale combo
'PARAMETRO 1    : oggetto combo box da caricare
'PARAMETRO 2    : codice articolo
'PARAMETR0 3    : true per caricare i dati del combo; false per assegnare solo il valore
'PARAMETRO 4    : valore da assegnare al combo
Sub CaricaComboRaggruppaProd(ByVal objCombo As ComboBox, _
                                ByVal strCodArt As String, _
                                bolCarica As Boolean, _
                                strValSalva As String, _
                                Optional vntValCombo As Variant)
Dim bolEnd As Boolean
Dim intAus As Integer
Dim strSQL As String
Dim hSS As CRecordSet

    If (bolCarica) Then
        'carico i valori del combo
        intAus = InStr(strCodArt, MXNU.SepVar)
        If (intAus = 0) Then intAus = Len(strCodArt) + 1
        strSQL = "SELECT CodTipologia,NumeroTip" _
                & " FROM TipologieArticoli" _
                & " WHERE CodiceArt=" & hndDBArchivi.FormatoSQL(Left$(strCodArt, intAus - 1), DB_TEXT) _
                & " ORDER BY NumeroTip"
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        Call objCombo.Clear
        Call objCombo.addItem("")
        Call objCombo.addItem(MXNU.CaricaStringaRes(75058))
        strValSalva = " R"
        bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
        Do While (Not bolEnd)
            Call objCombo.addItem(MXNU.CaricaStringaRes(75059) & " " & MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "CodTipologia", ""))
            strValSalva = strValSalva & CStr(MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "NumeroTip", 0))
            bolEnd = (Not MXDB.dbSuccessivo(hSS, TIPO_SNAPSHOT))
        Loop
        Call MXDB.dbChiudiSS(hSS)
    End If
    'assegna il valore al combo
    If Not IsMissing(vntValCombo) Then
        If (Trim$(vntValCombo) = " ") Then
            intAus = 0
        ElseIf (vntValCombo = "R") Then
            intAus = 1
        Else
            intAus = 2 + (Asc(vntValCombo) - 49)
        End If
        If (intAus < 0) Then
            intAus = 0
        ElseIf (intAus > objCombo.ListCount) Then
            intAus = objCombo.ListCount
        End If
        If (objCombo.ListCount > 0) Then objCombo.listIndex = intAus
    End If
End Sub

'NOME           : MovimentaArticolo
'DESCRIZIONE    : restituisce il flag di movimentazione dell'articolo
'PARAMETRO 1    : codice articolo
Function MovimentaArticolo(ByVal strCodArt As String) As Boolean
Dim strSQL As String
Dim hSS As CRecordSet

    strSQL = "SELECT AggiornaMag" _
            & " FROM AnagraficaArticoli" _
            & " WHERE Codice = " & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    MovimentaArticolo = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "AggiornaMag", True)
    Call MXDB.dbChiudiSS(hSS)
End Function

Public Function LeggiListinoArticolo(ByVal strArt As String, ByVal lngListino, Prezzo As Variant, PrezzoEuro As Variant)
    
    Dim hSS As MXKit.CRecordSet, q
    
    With MXDB
        Set hSS = .dbCreaSS(hndDBArchivi, "SELECT CODART, NRLISTINO,PREZZO,PREZZOEURO FROM LISTINIARTICOLI WHERE CODART=" & hndDBArchivi.FormatoSQL(strArt, DB_TEXT) & " AND NRLISTINO =" & lngListino)
        If .dbFineTab(hSS) Then
            Prezzo = 0
            PrezzoEuro = 0
            LeggiListinoArticolo = False
        Else
            Prezzo = .dbGetCampo(hSS, TIPO_SNAPSHOT, "PREZZO", 0)
            PrezzoEuro = .dbGetCampo(hSS, TIPO_SNAPSHOT, "PREZZOEURO", 0)
            LeggiListinoArticolo = True
        End If
        q = .dbChiudiSS(hSS)
    End With
    
End Function

Function LeggiContropartitaArticolo(ByVal CodArt As String, ByVal NrControPCont As Long, ByVal TipoConto As String, ByVal Nazione As Long) As String
    
    Dim Found As Boolean
    Dim hSS As MXKit.CRecordSet
    Dim coll As Collection
    Dim res As String
    
    Found = False
    res = ""
    If NrControPCont <> 0 Then
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT CodArt,Numero,SCGen FROM ControPartArticoli WHERE CodArt=" & _
            hndDBArchivi.FormatoSQL(CodArt, DB_TEXT) & " AND Esercizio = " & MXNU.AnnoAttivo & " AND Numero=" & NrControPCont)
        
        Found = Not MXDB.dbFineTab(hSS)
        If Found Then
            res = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "SCGen", "")
        End If
        Call MXDB.dbChiudiSS(hSS)
    End If
    If Not Found Then
        Set coll = New Collection
        If CaricaDatiArticolo(CodArt, "SCGenVenditeIta,SCGenVenditeEst,SCGenAcquistiIta,SCGenAcquistiEst", artCommerciali, coll) Then
            If TipoConto = "C" Then
                If Nazione = 0 Then
                    res = coll("SCGenVenditeIta")
                Else
                    res = coll("SCGenVenditeEst")
                End If
            ElseIf TipoConto = "F" Then
                If Nazione = 0 Then
                    res = coll("SCGenAcquistiIta")
                Else
                    res = coll("SCGenAcquistiEst")
                End If
            Else
                res = coll("SCGenVenditeIta")
            End If
        End If
        Set coll = Nothing
    End If
    LeggiContropartitaArticolo = res
End Function

Sub ScomponiCodiceArticolo(ByVal strArticolo As String, _
    Optional strCodiceNeutro As String, _
    Optional intPosSeparatore As Integer, _
    Optional strVarianti As String, _
    Optional bolAVarianti As Boolean)

    intPosSeparatore = InStr(strArticolo, MXNU.SepVar)
    bolAVarianti = (intPosSeparatore <> 0)
    If (bolAVarianti) Then
        strCodiceNeutro = Left$(strArticolo, intPosSeparatore - 1)
        strVarianti = Mid$(strArticolo, intPosSeparatore + 1)
    Else
        strCodiceNeutro = strArticolo
        strVarianti = ""
    End If
End Sub

Function VincolaUM(varListino As Variant, Optional bolListinoTrasformazione As Boolean = False) As Boolean
    Dim hSS As MXKit.CRecordSet
    Dim intq As Integer
    Dim strSQL As String
    
    If bolListinoTrasformazione Then
        strSQL = "SELECT VincolaUM FROM TabListiniTrasformazione WHERE NrListino=" & varListino
    Else
        strSQL = "SELECT VincolaUM FROM TabListini WHERE NrListino=" & varListino
    End If
    
    VincolaUM = False
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL)
    If MXDB.dbGetCampo(hSS, NO_REPOSITION, "VincolaUM", 0) = 1 Then
        VincolaUM = True
    End If
    intq = MXDB.dbChiudiSS(hSS)
    
End Function

Public Function LeggiVariantiArticolo(ByVal strCodArt As String, colPar As Collection) As Boolean
Dim cArt As MXBusiness.CVArt
Dim vntTipVar As Variant
Dim strVar As String
Dim inti As Integer

    Set cArt = MXART.CreaCVArt()
    LeggiVariantiArticolo = cArt.Valida(CHIEDIVAR_NESSUNA, False, strCodArt)
    If LeggiVariantiArticolo Then
        LeggiVariantiArticolo = (Len(cArt.VariantiEsplicite) > 0)
        If LeggiVariantiArticolo Then
            For Each vntTipVar In Split(Left$(cArt.VariantiEsplicite, Len(cArt.VariantiEsplicite) - 1), ";")
                strVar = Split(vntTipVar, "=")(1)
                inti = inti + 1
                colPar.Add strVar, CStr(inti)
            Next vntTipVar
        End If
    End If
    Call cArt.Termina
    Set cArt = Nothing
    
End Function

'------------------------------------------------------------
'nome:          Data2Esercizio
'descrizione:   restituisce l'esercizio di pertinenza della data passata
'parametri:     (in) data da controllare
'               (out) esercizio
'ritorno:       esito dell'operazione; se false la data è fuori
'               da tutti gli esercizi attualmente inseriti in tabella e viene restituito 0
'annotazioni:   rif.A-4767
'------------------------------------------------------------
Public Function i_Data2Esercizio(ByVal dteData As Variant, intEsercizio As Integer) As Boolean
Dim bolRes As Boolean
Dim strQuery As String
Dim hRSData As MXKit.CRecordSet

    bolRes = True
    On Local Error GoTo Data2Esercizio_ERR
    intEsercizio = 0
    With MXDB
        strQuery = "select CODICE" _
            & " from TABESERCIZI" _
            & " where DATAINIMAG<=" & hndDBArchivi.FormatoSQL(dteData, DB_DATE) _
            & " and DATAFINEMAG>=" & hndDBArchivi.FormatoSQL(dteData, DB_DATE)
        Set hRSData = .dbCreaSS(hndDBArchivi, strQuery)
        bolRes = Not .dbFineTab(hRSData)
        If (bolRes) Then
            intEsercizio = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "CODICE", 0)
            bolRes = (intEsercizio <> 0)
        End If
    End With
    
Data2Esercizio_END:
    Call MXDB.dbChiudiSS(hRSData)
    i_Data2Esercizio = bolRes
    On Local Error GoTo 0
    Exit Function

Data2Esercizio_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("Data2Esercizio", lngErrCod, strErrDsc))
    Resume Data2Esercizio_END
End Function


'------------------------------------------------------------
'nome:          LeggiDatiRiordino
'descrizione:   lettura dei dati di riordino di un dato articolo
'parametri:     Articolo: (in) codice articolo da considerare
'               Fornitore: (in/out) se passato cerca il fornitore fra quelli preferenziale e alternativi
'               Provenienza: (out) restituisce la provenienza dell'articolo
'               GGApprontamento: (out) restituisce i giorni di approntamento dell'articolo
'               GGApprovvigionamento: (out) restituisce i giornii di approvvigionamento dell'articolo
'               LottoRiferimento: (out) restituisce il lotto di riferimento per il tempo di approvvigionamento dell'articolo
'               TipoArrotondamento: (out) restituisce la modalità di arrotondamento del tempo di approvvigionamento rispetto al lotto
'ritorno:       esito dell'operazione
'annotazioni:   rif.A#5292
'------------------------------------------------------------
Public Function LeggiDatiRiordino(ByVal Articolo As String, ByRef fornitore As String, _
    Optional ByRef Provenienza As setProvenienzaArticolo, _
    Optional ByRef GGApprontamento As Long, _
    Optional ByRef GGApprovvigionamento As Long, _
    Optional ByRef LottoRiferimento As Variant, _
    Optional ByRef UmLottoRif As String, _
    Optional ByRef TipoArrotondamento As setArrotondaLeadTime) As Boolean
    
Dim bolRes As Boolean
Dim strQuery As String
Dim hRSData As MXKit.CRecordSet
Dim strFornitoreIn As String
Dim strSuffisso As String
Dim sCodicePadre As String
Dim sVarianti As String
Dim bArtVarianti As Boolean
Const SEGNAPOSTO_ARTICOLO = "%ARTICOLO%"

    bolRes = True
    On Local Error GoTo LeggiDatiRiordino_ERR
    strFornitoreIn = fornitore
    
    If (Len(Articolo) = 0) Then
        fornitore = ""
        GGApprontamento = 0
        GGApprovvigionamento = 0
        LottoRiferimento = CDec(0)
        UmLottoRif = ""
        TipoArrotondamento = arrLTProporzionale
    Else
        'RIF.A#6559 - determino il codice padre
        Call SeparaVarianti_i(Articolo, sCodicePadre, sVarianti)
        bArtVarianti = (Articolo <> sCodicePadre)
        
        With MXDB
            'determino i dati dell'articolo
            strQuery = "select PROVENIENZA," _
                & "FORNPREFACQ,TAPPRONTACQ,TAPPROVVACQ,LOTTORIFACQ,UMLOTTOACQ,ARROTLOTTOACQ," _
                & "(select top 1 CODFOR from TABVINCOLIPRODUZIONE) as FORNPREFPROD,TAPPRONTPROD,TAPPROVVPROD,LOTTORIFPROD,UMLOTTOPROD,ARROTLOTTOPROD," _
                & "FORNPREFLAV,TAPPRONTLAV,TAPPROVVLAV,LOTTORIFLAV,UMLOTTOLAV,ARROTLOTTOLAV" _
                & " from ANAGRAFICAARTICOLIPROD" _
                & " where CODICEART=" & SEGNAPOSTO_ARTICOLO & " and ESERCIZIO=" & MXNU.AnnoAttivo
            Set hRSData = .dbCreaSS(hndDBArchivi, Replace(strQuery, SEGNAPOSTO_ARTICOLO, hndDBArchivi.FormatoSQL(Articolo, DB_TEXT)))    'RIF.A#6559 - prima lettura: codice articolo
            'RIF.A#6559 - lettura dati dal padre se codice non generato
            If (.dbFineTab(hRSData) And bArtVarianti) Then
                Call .dbChiudiSS(hRSData)
                Set hRSData = .dbCreaSS(hndDBArchivi, Replace(strQuery, SEGNAPOSTO_ARTICOLO, hndDBArchivi.FormatoSQL(sCodicePadre, DB_TEXT)))
            End If
            If (.dbFineTab(hRSData)) Then
                bolRes = False
                GoTo LeggiDatiRiordino_END
            Else
                Provenienza = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "PROVENIENZA", PA_daAcquisto)
                Select Case Provenienza
                    Case PA_daAcquisto: strSuffisso = "ACQ"
                    Case PA_daProduzione: strSuffisso = "PROD"
                    Case PA_daContoLavoro: strSuffisso = "LAV"
                End Select
                fornitore = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "FORNPREF" & strSuffisso, "")
                bolRes = (Len(fornitore) > 0)
                'confronto con il fornitore passato
                If (bolRes And Len(strFornitoreIn) > 0) Then
                    bolRes = (fornitore = strFornitoreIn)
                End If
                'se trovato il fornitore principale => leggo i dati
                'RIF.A#9901 - i dati di riordino generali devono essere letti indipenentemente dalla presenza del fornitore preferenziale
                'If (bolRes) Then
                    GGApprontamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "TAPPRONT" & strSuffisso, 0)
                    GGApprovvigionamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "TAPPROVV" & strSuffisso, 0)
                    LottoRiferimento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "LOTTORIF" & strSuffisso, 0)
                    UmLottoRif = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "UMLOTTO" & strSuffisso, 0)
                    TipoArrotondamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "ARROTLOTTO" & strSuffisso, 0)
                'End If
            End If
            Call .dbChiudiSS(hRSData)
            'se non trovato il fornitore principale => leggo i dati dei fornitori alternativi
            If (Not bolRes) Then
                strQuery = "select top 1 CODFOR,GGAPPRONT,GGAPPROVV,LOTTORIF,UM,ARROTLOTTO" _
                    & " from TABLOTTIRIORDINO" _
                    & " where CODART=" & SEGNAPOSTO_ARTICOLO _
                    & " and TIPORIORD=" & Provenienza
                If (Len(strFornitoreIn) > 0) Then
                    strQuery = strQuery & " and CODFOR=" & hndDBArchivi.FormatoSQL(strFornitoreIn, DB_TEXT)
                End If
                'RIF.A#8830 - l'ordinamento va fatto per percentuale di ripartizione e, a parità di percentuale per posizione
                strQuery = strQuery & " order by PRCRIPART desc, NUMERO asc"
                
                Set hRSData = .dbCreaSS(hndDBArchivi, Replace(strQuery, SEGNAPOSTO_ARTICOLO, hndDBArchivi.FormatoSQL(Articolo, DB_TEXT))) 'RIF.A#6559 - prima lettura: codice articolo
                'RIF.A#6559 - lettura dati dal padre se codice non generato
                If (.dbFineTab(hRSData) And bArtVarianti) Then
                    Call .dbChiudiSS(hRSData)
                    Set hRSData = .dbCreaSS(hndDBArchivi, Replace(strQuery, SEGNAPOSTO_ARTICOLO, hndDBArchivi.FormatoSQL(sCodicePadre, DB_TEXT)))
                End If
                If (.dbFineTab(hRSData)) Then
                    bolRes = False
                    GoTo LeggiDatiRiordino_END
                Else
                    fornitore = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "CODFOR", 0)
                    GGApprontamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "GGAPPRONT", 0)
                    GGApprovvigionamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "GGAPPROVV", 0)
                    LottoRiferimento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "LOTTORIF", 0)
                    UmLottoRif = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "UM", 0)
                    TipoArrotondamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "ARROTLOTTO", 0)
                End If
                Call .dbChiudiSS(hRSData)
            End If
        End With
    End If
    
LeggiDatiRiordino_END:
    Call MXDB.dbChiudiSS(hRSData)
    LeggiDatiRiordino = bolRes
    On Local Error GoTo 0
    Exit Function

LeggiDatiRiordino_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("LeggiDatiRiordino", lngErrCod, strErrDsc))
    Resume LeggiDatiRiordino_END
End Function

'------------------------------------------------------------
'nome:          i_GeneraPartita
'descrizione:   genera una partita se non presente
'parametri:     codice articolo
'               codice partita
'ritorno:       esito dell'operazione
'annotazioni:
'------------------------------------------------------------
Public Function i_GeneraPartita(ByVal strArticolo As String, ByVal strPartita As String) As Boolean
Dim bolRes As Boolean
Dim hRSPart As MXKit.CRecordSet
Dim hRSCar As MXKit.CRecordSet
Dim strSQL As String
Dim lngNrRiga As Long
Dim vntValue As Variant

    bolRes = True
    On Local Error GoTo GeneraPartita_ERR
    If ((Len(strArticolo) > 0) And (Len(strPartita) > 0)) Then
        With MXDB
            'controlla se la partita è già stata generata
            strSQL = "select CODLOTTO" _
                & " from ANAGRAFICALOTTI" _
                & " where CODARTICOLO=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) _
                & " and CODLOTTO=" & hndDBArchivi.FormatoSQL(strPartita, DB_TEXT)
            Set hRSPart = .dbCreaSS(hndDBArchivi, strSQL)
            If (.dbFineTab(hRSPart)) Then
                'se la partita non c'è la genero
                strSQL = "insert into ANAGRAFICALOTTI (CODARTICOLO,CODLOTTO,BLOCCATO,UTENTEMODIFICA,DATAMODIFICA)" _
                    & " values (" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(strPartita, DB_TEXT) & ",0," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(Now, DB_DATETIME) & ")"
                Call .dbEseguiSQL(hndDBArchivi, strSQL)
    '            Call .dbInserisci(hDYPart)
    '            Call .dbSetCampo(hDYPart, TIPO_DYNASET, "CODARTICOLO", strArticolo)
    '            Call .dbSetCampo(hDYPart, TIPO_DYNASET, "CODLOTTO", strPartita)
    '            Call .dbSetCampo(hDYPart, TIPO_DYNASET, "BLOCCATO", CStr(vbUnchecked))
    '            Call .dbRegistra(hDYPart)
                'genero le caratteristiche in base al default
                strSQL = "select NRRIGA,CARATTDEFAULT" _
                    & " from TABCARATTLOTTI" _
                    & " where CODICE=(select CATEGORIA from ANAGRAFICAARTICOLI where CODICE=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) & ")"
                Set hRSCar = .dbCreaSS(hndDBArchivi, strSQL)
                Do While Not (.dbFineTab(hRSCar))
                    lngNrRiga = .dbGetCampo(hRSCar, TIPO_SNAPSHOT, "NRRIGA", 0)
                    vntValue = .dbGetCampo(hRSCar, TIPO_SNAPSHOT, "CARATTDEFAULT", "")
                    strSQL = "insert into ANAGRCARLOTTI (CODARTICOLO,CODLOTTO,NRRIGA,VALORE,UTENTEMODIFICA,DATAMODIFICA)" _
                        & " values (" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(strPartita, DB_TEXT) _
                        & "," & lngNrRiga & "," & hndDBArchivi.FormatoSQL(vntValue, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(Now, DB_DATETIME) & ")"
                    Call .dbEseguiSQL(hndDBArchivi, strSQL)
                    
                    Call .dbSuccessivo(hRSCar)
                Loop
                Call .dbChiudiSS(hRSPart)
            End If
        End With
    End If
    
GeneraPartita_END:
    Call MXDB.dbChiudiSS(hRSCar)
    Call MXDB.dbChiudiDY(hRSPart)
    i_GeneraPartita = bolRes
    On Local Error GoTo 0
    Exit Function

GeneraPartita_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("GeneraPartita", lngErrCod, strErrDsc))
    Resume GeneraPartita_END
Resume
End Function

'------------------------------------------------------------
'nome:          LeggiFlagGeneraPartita
'descrizione:   legge per l'anno attivo il valore del flagPartita dalla tabella TABVINCOLIGIC ovvero l'option button Partita dalla form Vincoli
'parametri:
'ritorno:       valore del flag per l'esercizio attivo
'annotazioni:   RIF.A#5948
'------------------------------------------------------------
Public Function LeggiFlagGeneraPartita() As setFlagGeneraPartita
    Dim strQuery As String
    Dim hRS As MXKit.CRecordSet
    
    On Local Error GoTo ERR_LeggiFlagGeneraPartita
        
    strQuery = "select FLGPARTITA" _
        & " from TABVINCOLIGIC" _
        & " where ESERCIZIO = " & MXNU.AnnoAttivo
    With MXDB
        Set hRS = .dbCreaSS(hndDBArchivi, strQuery)
        LeggiFlagGeneraPartita = .dbGetCampo(hRS, TIPO_SNAPSHOT, "FLGPARTITA", 0)
    End With

END_LeggiFlagGeneraPartita:
    Call MXDB.dbChiudiSS(hRS)
    Set hRS = Nothing
    On Local Error GoTo 0
    Exit Function

ERR_LeggiFlagGeneraPartita:
    Dim lngErrCod As Long
    Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("LeggiFlagGeneraPartita", lngErrCod, strErrDsc))
    Resume END_LeggiFlagGeneraPartita
    
End Function

Public Function SeparaVarianti_i(ByVal strCod As String, strArtTip As String, strVar As String) As Boolean
    
    Dim psep As Integer
    
    psep = InStr(strCod, MXNU.SepVar)
    SeparaVarianti_i = psep > 0
    If psep > 0 Then
        strArtTip = Left$(strCod, psep - 1)
        strVar = Mid$(strCod, psep + 1)
    Else
        strArtTip = ""
        strVar = ""
    End If
End Function

'RIF.A#6234 - restituisce un valore booleano che indica se l'articolo movimenta o meno le matricole
Public Function ArticoloMovimentaMatricole(ByVal strArticolo As String) As Boolean
Dim strQuery As String
Dim hRSData As MXKit.CRecordSet

    With MXDB
        strQuery = "select MOVIMENTAMATRICOLE from ANAGRAFICAARTICOLI where CODICE=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
        Set hRSData = .dbCreaSS(hndDBArchivi, strQuery)
        ArticoloMovimentaMatricole = (.dbGetCampo(hRSData, TIPO_SNAPSHOT, "MOVIMENTAMATRICOLE", 0) <> 0)
        Call .dbChiudiSS(hRSData)
    End With
End Function

Public Function ArticoloFloorStock(ByVal strArticolo As String) As Boolean
Dim strQuery As String
Dim hRSData As MXKit.CRecordSet
Dim strCodicePadre As String

    With MXDB
        strQuery = "select FLOORSTOCK" _
            & " from ANAGRAFICAARTICOLIPROD" _
            & " where CODICEART=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) _
            & " and ESERCIZIO=" & MXNU.AnnoAttivo
        Set hRSData = .dbCreaSS(hndDBArchivi, strQuery)
        If (.dbFineTab(hRSData)) Then
            'se articolo a varianti non generato => leggo il dato dall'articolo padre
            Call .dbChiudiSS(hRSData)
            Call ScomponiCodiceArticolo(strArticolo, strCodicePadre)
            If (strCodicePadre <> strArticolo) Then
                strQuery = "select FLOORSTOCK" _
                    & " from ANAGRAFICAARTICOLIPROD" _
                    & " where CODICEART=" & hndDBArchivi.FormatoSQL(strCodicePadre, DB_TEXT) _
                    & " and ESERCIZIO=" & MXNU.AnnoAttivo
                Set hRSData = .dbCreaSS(hndDBArchivi, strQuery)
            End If
        End If
        'leggo e restituisco il risultato
        If (.dbFineTab(hRSData)) Then
            ArticoloFloorStock = False
        Else
            ArticoloFloorStock = (.dbGetCampo(hRSData, TIPO_SNAPSHOT, "FLOORSTOCK", 0) <> 0)
        End If
        
        Call .dbChiudiSS(hRSData)
    End With
End Function
