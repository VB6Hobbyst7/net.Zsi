Attribute VB_Name = "MAgenti"
Option Explicit
DefLng A-Z
Public Function EstensioniRegoleBusiness(agt As MXKit.CAgenteAuto, strComando As String, parms As String)

Dim ArtComposto As MXBusiness.CComposto
Dim ArtComponente As MXBusiness.CComponente
Dim ArtLavorato As MXBusiness.CCicloLav
Dim FasiLav As MXBusiness.CFaseLav

Dim strVer As String
Dim strDat As String
Dim retVal As String
Dim strArt As String
Dim filNom As String
Dim filNum As Integer
Dim filErr As Integer
Dim strOut As String
Dim strPar As String
Dim NumCmp As Integer
Dim curCmp As Integer
Dim parTpS As String
Dim parQta As Variant
Dim parMLv As Integer
Dim parTpF As Integer
Dim strBuf As String
Dim parArt As String
Dim parVer As String
Dim parDat As String
Dim strUM As String
Dim parUM As String
Const ESP_POS_ART = 1
Const ESP_POS_VER = 2
Const ESP_POS_DAT = 3
Const ESP_POS_QTA = 4
Const ESP_POS_UM = 10
Const ARC_POS_ART = 1
Const ARC_POS_VER = 2
Const ARC_POS_DAT = 3
Const ARC_POS_RET = 4
Const ARC_POS_UM = 5
            
Select Case strComando
    
    Case "#ESPLODIDISTINTA"
        'lunghezza campi per file di testo
        Const LEN_ART = 50
        Const LEN_TCP = 1
        Const LEN_PZL = 19
        Const LEN_PZE = 19
        Const LEN_LTC = 5
        Const LEN_DSC = 80
        Const LEN_VER = 10
        Const LEN_DAT = 10
        Const LEN_QTA = 16
        Const LEN_LIV = 5
        Const LEN_LVP = 1
        Const LEN_OPR = 1
        Const LEN_UM = 3    ' Rif. anomalia n.ro 4947 (prima della correzione la costante era pari a 2)
        Const LEN_QTARIB = 16
        '*** parametri della funzione ***
        Const ESP_POS_LIV = 5
        Const ESP_POS_FIL = 6
        Const ESP_POS_TPF = 7 'tipo campi (0 = lunghezza fissa o 1=lunghezza variabile)
        Const ESP_POS_VLC = 8   'tipo valorizzazione costi:
                    'UPA    = Valorizzazione all'Ultimo Prezzo di Acquisto 0
                    'CMD    = Valorizzazione al Costo Medio 2
                    'LIFO   = Valorizzazione al LIFO 3
                    'LSTn   = Valorizzazione al Listino n 5
                    'EXTN   = Valorizzazioen a Campo Extra n 4
        Const ESP_POS_TPS = 9 'tipo di separatore per campi a lunghezza variabile
        'rif.S-1105 - gestione del tipo di elaborazione
        'tipo di valutazioni:
        '   NOVAL:  non effettua alcuna valutazione aggiuntiva
        '   DISP:   effettua la valutazione delle disponibilità
        '   COSTI:  effettua la valorizzazione
        '   DATE:   effettua la valutazione delle date richiesta
        '   MSG:    messaggio per componenti non validi
        '   LOOP:   controlla e blocca eventuali loop di struttura
        'RIF.S#906
        '   TECN:   applica il criterio tecnico per l'esplosione dei componenti (attivo per default)
        '   GEST:   applica il criterio gestionale per l'esplosione dei componenti
        Const ESP_POS_TIPO = 11 'tipo di elaborazione
        Const ESP_POS_DEPOSITI = 12 'RIF.S#649 - depositi da considerare
        Const ESP_POS_PARTITA = 13 'RIF.S#739 - partita di materiale
        
        Dim enmTipoValutazione As MXBusiness.setValutaDistinta
        Dim strConsideraDepositi As String 'RIF.S#649 - string depositi da considerare
        
        Set ArtComposto = MXDBA.CreaCComposto
        
        '*** estrae i parametri della funzione
        'codice articolo
        strPar$ = i2s(parms$, ESP_POS_ART, ";")
        parArt$ = agt.LeggiVariabile(strPar$)
        'versione distinta
        strPar$ = i2s(parms$, ESP_POS_VER, ";")
        parVer$ = agt.LeggiVariabile(strPar$)
        'data di valutazione
        strPar$ = i2s(parms$, ESP_POS_DAT, ";")
        parDat = agt.LeggiVariabile(strPar$)
        If (Not IsDate(parDat$)) Then parDat$ = Format$(Date, MXNU.Formato_Data)
        'quantità distinta
        strPar$ = i2s(parms$, ESP_POS_QTA, ";")
        strBuf$ = agt.LeggiVariabile(strPar$)
        'RIF.A#11119 - eliminato l'arrotondamento della quantità
        If (Trim$(strBuf$) = "") Then parQta = 0 Else parQta = Val(strBuf$)
        'UM
        strUM = i2s(parms$, ESP_POS_UM, ";")
        parUM = agt.LeggiVariabile(strUM)
        'massimo livello
        strPar$ = i2s(parms$, ESP_POS_LIV, ";")
        strBuf$ = agt.LeggiVariabile(strPar$)
        If (Trim$(strBuf$) = "") Then parMLv% = 0 Else parMLv% = Val(strBuf$)
        'tipo file
        strBuf$ = i2s(parms$, ESP_POS_TPF, ";")
        parTpF% = Val(strBuf$)
        'tipo valorizzazione
        strBuf$ = Trim$(i2s(parms$, ESP_POS_VLC, ";"))
        If (strBuf$ = "UPA") Then
            ArtComposto.pTipoValorizza = valUltimoPrezzoAcquisto
        ElseIf (strBuf$ = "CMD") Then
            ArtComposto.pTipoValorizza = valCostoMedio
        ElseIf (strBuf$ = "LIFO") Then
            ArtComposto.pTipoValorizza = valLIFO
        ElseIf (Left$(strBuf$, 3) = "LST") Then
            ArtComposto.pTipoValorizza = valListino
            ArtComposto.pCampoValorizza = Val(Mid$(strBuf$, 4))
        ElseIf (Left$(strBuf$, 3) = "EXT") Then
            ArtComposto.pTipoValorizza = valExtra
            ArtComposto.pCampoValorizza = Val(Mid$(strBuf$, 4))
        ElseIf (strBuf$ = "UPS") Then 'RIF.S#1800
            ArtComposto.pTipoValorizza = valUltimoPrezzoSpese
        Else
            ArtComposto.pTipoValorizza = valCostoStandard
            ArtComposto.pCampoValorizza = Mid$(strBuf$, 4)
        End If
        'tipo separatore
        strBuf$ = i2s(parms$, ESP_POS_TPS, ";")
        If Left(strBuf$, 1) = Chr$(34) Then
           'stringhe delimitate dal carattere "
           parTpS = Mid(strBuf$, 2, Len(strBuf$) - 2)
        Else
           'stringhe contenute in variabili
           parTpS = agt.LeggiVariabile(strBuf$)
           If parTpS = "" Then parTpS = ","
        End If
        'rif.S-1105 - tipo di valutazione
        strBuf = i2s(parms, ESP_POS_TIPO, ";")
        If (Len(strBuf) = 0) Then
            'fa come prima
            enmTipoValutazione = flgLivProduzione + flgTipoComposto + flgDisponibilita + flgValorizza
        Else
            Dim vntDummy As Variant
            enmTipoValutazione = flgLivProduzione + flgTipoComposto
            For Each vntDummy In Split(strBuf, "|")
                Select Case UCase$(vntDummy)
                    Case "NOVAL"
                        enmTipoValutazione = flgLivProduzione + flgTipoComposto
                    Case "DISP"
                        enmTipoValutazione = enmTipoValutazione + flgDisponibilita
                        'RIF.S#649 - leggo la stringa dei depositi da considerare
                        strBuf = i2s(parms, ESP_POS_DEPOSITI, ";")
                        ArtComposto.ConsideraDepositi = Replace(strBuf, "|", vbNullChar)
                    Case "COSTI"
                        enmTipoValutazione = enmTipoValutazione + flgValorizza
                    Case "DATE"
                        enmTipoValutazione = enmTipoValutazione + flgValutaDate
                    Case "MSG"
                        enmTipoValutazione = enmTipoValutazione + flgMessaggioSeNonValido
                    Case "LOOP"
                        enmTipoValutazione = enmTipoValutazione + flgCtrlLoop + flgBloccaLoop
                    Case "TECN"
                        enmTipoValutazione = enmTipoValutazione And (Not flgBloccaSeAcquisto)
                    Case "GEST"
                        enmTipoValutazione = enmTipoValutazione Or flgBloccaSeAcquisto
                End Select
            Next vntDummy
        End If
        'RIF.S#739 - partita di materiale
        strBuf = i2s(parms, ESP_POS_PARTITA, ";")
        If (Len(strBuf) > 0) Then
            ArtComposto.pPartitaAssegnata = agt.LeggiVariabile(strBuf)
        End If
        
        'effettua l'esplosione della distinta
        If ArtComposto.EsplodiDistinta(parArt$, parVer$, parDat$, parQta, parUM, parDat$, parMLv%, enmTipoValutazione) Then
               
            '*** scrive i risultati su un file di testo
            filNom$ = MXNU.GetTempFile()
            GoSub Aprifile
            'prima riga: dati della testa
            strOut$ = ""
            If (parTpF% = 0) Then
                'campi lunghezza fissa
                strOut$ = strOut$ & Right$(Space$(LEN_ART) & ArtComposto.pCodice, LEN_ART) & _
                     Right$(Space$(LEN_VER) & ArtComposto.pVersione, LEN_VER) & _
                     Right$(Space$(LEN_DAT) & ArtComposto.pDataValutazione, LEN_DAT) & _
                     Right$(Space$(LEN_QTA) & Str(ArtComposto.pQuantita), LEN_QTA) & _
                     Right$(Space$(LEN_LIV) & ArtComposto.pProgressivo, LEN_LIV) & _
                     Right$(Space$(LEN_UM) & ArtComposto.pUM, LEN_UM) ' ArtComposto.pProgressivo, LEN_UM)  <--(anomalia n.ro 6069)
            Else
                'campi lunghezza variabile
                strOut$ = strOut$ & ArtComposto.pCodice & parTpS & _
                    ArtComposto.pVersione & parTpS & ArtComposto.pDataValutazione & _
                    parTpS & Str(ArtComposto.pQuantita) & parTpS & Str(ArtComposto.pProgressivo) & parTpS & ArtComposto.pUM
            End If
            GoSub scriviRiga
            
            '************* Rif. scheda 2509 ***********************
            Dim colListaComponenti As Collection
            Set colListaComponenti = New Collection
            Call RiordinaCollectionDBA1(ArtComposto, colListaComponenti)
            
            'altre righe: dati della distinta esplosa
            For Each ArtComponente In colListaComponenti
                strOut$ = ""
                If (parTpF% = 0) Then
                    'campi lunghezza fissa
                    
                    ' Rif. anomalia n.ro 4799
                    Dim strCostoTotLire As String
                    Dim strCostoTotEuro As String
                    strCostoTotLire = Left$(ArtComponente.pCostoTotaleMateriale(tcsLire), LEN_PZL)
                    strCostoTotEuro = Left$(ArtComponente.pCostoTotaleMateriale(tcsEuro), LEN_PZE)
                    ' Fine rif. anomalia n.ro 4799
                    
                    ' Rif. anomalia #8486
                    Dim strQta1 As String
                    Dim strQta2 As String
                    strQta1 = Left$(ArtComponente.pQuantita1, LEN_QTA)
                    strQta2 = Left$(ArtComponente.pQuantita2, LEN_QTA)
                    ' Fine rif. anomalia #8486
                    
                    ' rif. anomalia n.ro 6047 (modificato  ArtComponente.pVersione, LEN_DSC con  ArtComponente.pVersione, LEN_VER)
                    strOut$ = strOut$ & Right$(Space$(LEN_LIV) & ArtComponente.pNrRiga, LEN_LIV) & _
                        Right$(Space$(LEN_LIV) & ArtComponente.pLivello, LEN_LIV) & _
                        Right$(Space$(LEN_LVP) & ArtComponente.pLivProduzione, LEN_LVP) & _
                        Right$(Space$(LEN_TCP) & ArtComponente.pTipoComponente, LEN_TCP) & _
                        Right$(Space$(LEN_ART) & ArtComponente.pCodice, LEN_ART) & _
                        Right$(Space$(LEN_DSC) & swapp(ArtComponente.pDescrizione, ",", " "), LEN_DSC) & _
                        Right$(Space$(LEN_VER) & ArtComponente.pVersione, LEN_VER) & _
                        Right$(Space$(LEN_QTA) & Str(strQta1), LEN_QTA) & _
                        Right$(Space$(LEN_OPR) & ArtComponente.pOperatore, LEN_OPR) & _
                        Right$(Space$(LEN_QTA) & Str(strQta2), LEN_QTA) & _
                        Right$(Space$(LEN_QTA) & Str(ArtComponente.pQtaComponente), LEN_QTA) & _
                        Right$(Space$(LEN_PZL) & Str(strCostoTotLire), LEN_PZL) & _
                        Right$(Space$(LEN_PZE) & Str(strCostoTotEuro), LEN_PZE) & _
                        Right$(Space$(LEN_QTA) & Str(ArtComponente.pGiacenza), LEN_QTA) & _
                        Right$(Space$(LEN_QTA) & Str(ArtComponente.pImpegnato(ArtComponente.pUM)), LEN_QTA) & _
                        Right$(Space$(LEN_QTA) & Str(ArtComponente.pOrdinato), LEN_QTA) & _
                        Right$(Space$(LEN_QTA) & Str(ArtComponente.pDisponibile), LEN_QTA) & _
                        Right$(Space$(LEN_QTA) & Str(ArtComponente.pResiduo), LEN_QTA) & _
                        Right$(Space$(LEN_LTC) & Str(ArtComponente.pLeadTimeGlobale), LEN_LTC) & _
                        Right$(Space$(LEN_DAT) & ArtComponente.pDataDisponibilita, LEN_DAT) & _
                        Right$(Space$(LEN_DAT) & ArtComponente.pDataRichiesta, LEN_DAT) & _
                        Right$(Space$(LEN_UM) & ArtComponente.pUM, LEN_UM) & _
                        Right$(Space$(LEN_QTARIB) & Str(ArtComponente.pQtaRibaltamentoCosti), LEN_QTARIB)
                Else
                    'campi lunghezza variabile
                    strOut$ = strOut$ & ArtComponente.pNrRiga & parTpS & _
                        ArtComponente.pLivello & parTpS & _
                        ArtComponente.pLivProduzione & parTpS & _
                        ArtComponente.pTipoComponente & parTpS & _
                        ArtComponente.pCodice & parTpS & _
                        swapp(ArtComponente.pDescrizione, ",", " ") & parTpS & _
                        ArtComponente.pVersione & parTpS & _
                        Str(ArtComponente.pQuantita1) & parTpS & _
                        ArtComponente.pOperatore & parTpS & _
                        Str(ArtComponente.pQuantita2) & parTpS & _
                        Str(ArtComponente.pQtaComponente) & parTpS & _
                        Str(ArtComponente.pCostoTotaleMateriale(tcsLire)) & parTpS & _
                        Str(ArtComponente.pCostoTotaleMateriale(tcsEuro)) & parTpS & _
                        Str(ArtComponente.pGiacenza) & parTpS & _
                        Str(ArtComponente.pImpegnato(ArtComponente.pUM)) & parTpS & _
                        Str(ArtComponente.pOrdinato) & parTpS & _
                        Str(ArtComponente.pDisponibile) & parTpS & _
                        Str(ArtComponente.pResiduo) & parTpS & _
                        ArtComponente.pLeadTimeGlobale & parTpS & _
                        ArtComponente.pDataDisponibilita & parTpS & _
                        ArtComponente.pDataRichiesta & parTpS & _
                        ArtComponente.pUM & parTpS & _
                        Str(ArtComponente.pQtaRibaltamentoCosti)
                End If
                GoSub scriviRiga
            Next
            GoSub chiudiFile
            
            Set colListaComponenti = Nothing 'Rif. scheda n.ro 2506
            
            'restituisce alla regola il nome del file temporaneo
            strBuf$ = i2s(parms$, ESP_POS_FIL, ";")
            Call agt.ScriviVariabile(strBuf$, filNom$)
        End If
        EstensioniRegoleBusiness = True
        Set ArtComponente = Nothing
        ArtComposto.Termina
        Set ArtComposto = Nothing
    Case "#ARTCOMPOSTO"
        Set ArtComposto = MXDBA.CreaCComposto

        'Codice ARTICOLO
        strArt$ = i2s(parms$, ARC_POS_ART, ";")
        parArt$ = agt.LeggiVariabile(strArt$)
        
        'versione distinta
        strVer$ = i2s(parms$, ARC_POS_VER, ";")
        parVer$ = agt.LeggiVariabile(strVer$)
        
        'data di valutazione
        strDat$ = i2s(parms$, ARC_POS_DAT, ";")
        parDat$ = agt.LeggiVariabile(strDat$)
        If (Not IsDate(parDat$)) Then parDat$ = Format$(Date, MXNU.Formato_Data)
        
        'leggi nome variabile UM
        strUM = i2s(parms$, ARC_POS_UM, ";")
        
        If ArtComposto.Valida(parArt$, False) Then
            If (ArtComposto.ArticoloComposto(ctrl_versione + ctrl_StatoVersione, parArt$, parVer$, parDat)) Then
                retVal$ = "0" 'articolo composto -> restituisco Versione e Data
                Call agt.ScriviVariabile(strArt$, ArtComposto.pCodice)
                Call agt.ScriviVariabile(strVer$, ArtComposto.pVersione)
                Call agt.ScriviVariabile(strDat$, ArtComposto.pDataValutazione)
                Call agt.ScriviVariabile(strUM, ArtComposto.pUM)
            Else
                retVal$ = "1" 'articolo non composto
            End If
        Else
        retVal$ = "2" 'articolo non valido
        End If
        strBuf$ = i2s(parms$, ARC_POS_RET, ";")
        Call agt.ScriviVariabile(strBuf$, retVal$)
        EstensioniRegoleBusiness = True
        ArtComposto.Termina
        Set ArtComposto = Nothing
    Case "#CERCACICLO"
        Set ArtLavorato = MXCICLI.CreaCCiclo()
        '*** estrae i parametri della funzione
        'codice articolo
        strPar$ = i2s(parms$, ARC_POS_ART, ";")
        parArt$ = agt.LeggiVariabile(strPar$)
        'versione distinta
        strVer$ = i2s(parms$, ARC_POS_VER, ";")
        parVer$ = agt.LeggiVariabile(strVer$)
        'data di valutazione
        strDat$ = i2s(parms$, ARC_POS_DAT, ";")
        parDat$ = agt.LeggiVariabile(strDat$)
        If (Not IsDate(parDat$)) Then parDat$ = Format$(Date, MXNU.Formato_Data)
        If ArtLavorato.Valida(parArt$, False) Then
            If ArtLavorato.ArticoloLavorato(ctrlCicloStatoVersione + ctrlCicloStatoVersione, parArt$, parVer$, parDat) Then
                retVal$ = "0" 'articolo composto -> restituisco Versione e Data
                Call agt.ScriviVariabile(strVer$, ArtLavorato.pVersione)
                Call agt.ScriviVariabile(strDat$, ArtLavorato.pDataValutazione)
            Else
                retVal$ = "1" 'articolo non collegato ad un ciclo
            End If
        Else
            retVal$ = "2" 'articolo non valido
        End If
        strBuf$ = i2s(parms$, ARC_POS_RET, ";")
        Call agt.ScriviVariabile(strBuf$, retVal$)
        EstensioniRegoleBusiness = True
        ArtLavorato.Termina
        Set ArtLavorato = Nothing
    Case "#VALUTACICLO"
        '*** parametri della funzione ***
        Const ESP_POS_DSCH = 5 'data richiesta per schedulazione
        Const ESP_POS_TIPS = 6
        Const ESP_POS_TSCH = 7 'tipo schedulazione
        Const ESP_POS_FILE = 8
        Const ESP_POS_UMCLV = 9
        
        Set ArtLavorato = MXCICLI.CreaCCiclo()
        Dim flagVal As Integer
        '*** strutture per la valorizzazione ***
        Dim vntTempo As Variant
        Dim cTime As MXBusiness.CCalcTime
        Dim enmUM As MXBusiness.setTimeUM
        
        Set cTime = New MXBusiness.CCalcTime
        '*** estrae i parametri della funzione
        'codice articolo
        strPar$ = i2s(parms$, ESP_POS_ART, ";")
        parArt$ = agt.LeggiVariabile(strPar$)
        'versione ciclo
        strPar$ = i2s(parms$, ESP_POS_VER, ";")
        parVer$ = agt.LeggiVariabile(strPar$)
        'data di valutazione
        strPar$ = i2s(parms$, ESP_POS_DAT, ";")
        parDat = agt.LeggiVariabile(strPar$)
        If (Not IsDate(parDat$)) Then parDat$ = Format$(Date, MXNU.Formato_Data)
        'quantità ciclo
        strPar$ = i2s(parms$, ESP_POS_QTA, ";")
        strBuf$ = agt.LeggiVariabile(strPar$)
        'RIF.A#11119 - tolto l'arrotondamento della quantità
        If (Trim$(strBuf$) = "") Then parQta = 0 Else parQta = Val(strBuf$)
        'flag durata
        '   data lavorazione
        Dim enmTipoSched As MXBusiness.setCicloLavSchedulazione
        strPar$ = i2s(parms$, ESP_POS_DSCH, ";")
        strBuf$ = agt.LeggiVariabile(strPar$)
        If Not IsDate(strBuf$) Then
            enmTipoSched = schLavNonSchedulare
            strBuf = Date 'rif.A-4496 - gestione schedulazione NODATE
        Else
            '   tipo schedulazione
            strPar$ = i2s(parms$, ESP_POS_TSCH, ";")
            If InStr(UCase$(strPar$), "INDIETRO") Then
                enmTipoSched = schLavSchedulaBackward
            Else
                enmTipoSched = schLavSchedulaForward
            End If
        End If
        Call ArtLavorato.ImpostaSchedulazione(enmTipoSched, CVDate(Format$(strBuf$, MXNU.Formato_Data)))
        'tipo separatore
        strBuf$ = i2s(parms$, ESP_POS_TIPS, ";")
        If Left(strBuf$, 1) = Chr$(34) Then
           'stringhe delimitate dal carattere "
           parTpS = Mid(strBuf$, 2, Len(strBuf$) - 2)
        Else
           'stringhe contenute in variabili
           parTpS = agt.LeggiVariabile(strBuf$)
        End If
        'unità di misura
        strPar$ = i2s(parms$, ESP_POS_UMCLV, ";")
        parUM = agt.LeggiVariabile(strPar$)
        If (Len(parUM) = 0) Then
            If ArtLavorato.ArticoloLavorato(ctrlCicloStatoVersione + ctrlCicloStatoVersione, parArt, parVer, parDat) Then
                parUM = ArtLavorato.pUM
            End If
        End If
        'gestione calcolo costi (rif.sch.3460)
        Call ArtLavorato.ImpostaCalcoloCosti(valCostiQuoteStandard)
        'effettua l'esplosione del ciclo
        'flagVal = VC_FLG_MESSAGGI + VC_FLG_NOSSE
        'rif.A#5134 - esclusione delle fasi non valide
        If ArtLavorato.ValutaCiclo(parArt$, parVer$, parDat, parQta, parUM, _
            flgCicloLavCtrlStatoVersione + tipoValCostiLavorazione + tipoValEscludiNonValide) Then
            '*** scrive i risultati su un file di testo
            filNom$ = MXNU.GetTempFile()
            GoSub Aprifile
            'prima riga: dati della testa
            strOut$ = ""
            'campi lunghezza variabile
            strOut$ = strOut$ & ArtLavorato.pCodice & parTpS & ArtLavorato.pVersione & parTpS _
                & ArtLavorato.pDataValutazione & parTpS & Str(ArtLavorato.pQuantita) & parTpS _
                & Str(ArtLavorato.pNumeroFasi)
            If enmTipoSched = schLavNonSchedulare Then
                strOut$ = strOut$ & parTpS & Now & parTpS & Now
            Else
                strOut$ = strOut$ & parTpS _
                & CVDate(ArtLavorato.pDataInizioLav & " " & ArtLavorato.pOraInizioLav) & parTpS _
                & CVDate(ArtLavorato.pDataFineLav & " " & ArtLavorato.pOraFineLav)
            End If
            GoSub scriviRiga
            'altre righe: dati del ciclo esploso
            For Each FasiLav In ArtLavorato.FasiLavorazione
                strOut$ = ""
                'campi lunghezza variabile
                strOut$ = strOut$ & FasiLav.pNumeroFase & parTpS _
                    & FasiLav.pTipoFase & parTpS _
                    & FasiLav.pIndiceMoltiplicatore & parTpS _
                    & FasiLav.pCodiceOperazione & parTpS _
                    & FasiLav.pCodiceCdLavoro & parTpS _
                    & FasiLav.pCodiceRisorsa & parTpS _
                    & Str(FasiLav.pQuantitaOrdinata) & parTpS
                'tempo attrezzaggio
                vntTempo = FasiLav.pTempoSetup
                enmUM = Val(FasiLav.pUMTempoSetup)
                GoSub InserisciTempo
                strOut$ = strOut$ & Str(FasiLav.pLottoTempoSetup) & parTpS
                'tempo risorsa
                vntTempo = FasiLav.pTempoRisorsa
                enmUM = Val(FasiLav.pUMTempoRisorsa)
                GoSub InserisciTempo
                strOut$ = strOut$ & Str(FasiLav.pLottoTempoRisorsa) & parTpS
                'tempo manodopera
                vntTempo = FasiLav.pTempoUomo
                enmUM = Val(FasiLav.pUMTempoUomo)
                GoSub InserisciTempo
                strOut$ = strOut$ & FasiLav.pLottoTempoUomo & parTpS
                'tempo totale
                vntTempo = FasiLav.pDurata
                enmUM = Val(FasiLav.pUMTempo(flTimeDurata))
                GoSub InserisciTempo
                strOut$ = strOut$ & FasiLav.pLottoTempo(flTimeDurata) & parTpS _
                    & Str(FasiLav.pTempoCoda(consTimePreventivo)) & parTpS _
                    & Str(FasiLav.pTempoMovimentazione(consTimePreventivo)) & parTpS _
                    & Str(FasiLav.pTempoAttraversamento) & parTpS _
                    & FasiLav.pDataInizioLav & parTpS _
                    & FasiLav.pOraInizioLav & parTpS _
                    & FasiLav.pDataFineLav & parTpS _
                    & FasiLav.pOraFineLav & parTpS
                'rif.A-3460 espongo i costi sia in lire e in euro
                strOut$ = strOut$ & Str(FasiLav.pCostoSetup(tcsLire)) & parTpS _
                    & Str(FasiLav.pCostoRisorsa(tcsLire)) & parTpS _
                    & Str(FasiLav.pCostoUomo(tcsLire)) & parTpS _
                    & Str(FasiLav.pCostoIndVariabile(tcsLire)) & parTpS _
                    & Str(FasiLav.pCostoIndFisso(tcsLire)) & parTpS _
                    & Str(FasiLav.pCostoInterno(tcsLire)) & parTpS _
                    & Str(FasiLav.pCostoEsterno(tcsLire)) & parTpS _
                    & Str(FasiLav.pCostoPieno(tcsLire)) & parTpS
                strOut$ = strOut$ & Str(FasiLav.pCostoSetup(tcsEuro)) & parTpS _
                    & Str(FasiLav.pCostoRisorsa(tcsEuro)) & parTpS _
                    & Str(FasiLav.pCostoUomo(tcsEuro)) & parTpS _
                    & Str(FasiLav.pCostoIndVariabile(tcsEuro)) & parTpS _
                    & Str(FasiLav.pCostoIndFisso(tcsEuro)) & parTpS _
                    & Str(FasiLav.pCostoInterno(tcsEuro)) & parTpS _
                    & Str(FasiLav.pCostoEsterno(tcsEuro)) & parTpS _
                    & Str(FasiLav.pCostoPieno(tcsEuro))
                GoSub scriviRiga
            Next
            GoSub chiudiFile
            'restituisce alla regola il nome del file temporaneo
            strBuf$ = i2s(parms$, ESP_POS_FILE, ";")
            Call agt.ScriviVariabile(strBuf$, filNom$)
        End If
        EstensioniRegoleBusiness = True
        Set cTime = Nothing
        Set FasiLav = Nothing
        ArtLavorato.Termina
        Set ArtLavorato = Nothing
    Case Else
        EstensioniRegoleBusiness = False
    End Select
    Exit Function

'subroutines per la gestione dei file per #eplodidistinta e #valutaciclo

'apre il file temporaneo
Aprifile:
    On Local Error Resume Next
    filNum% = FreeFile
    Open filNom$ For Output Access Write As filNum%
    On Local Error GoTo 0
Return

'scrive la stringa nel file temporaneo
scriviRiga:
    On Local Error Resume Next
    Print #filNum%, strOut$
    On Local Error GoTo 0
Return

'chiude il file temporaneo
chiudiFile:
    On Local Error Resume Next
    Close #filNum%
    On Local Error GoTo 0
Return

InserisciTempo:
    Select Case enmUM
    Case UM_TIME_ORE, UM_TIME_MINUTI
        'nessuna conversione
    Case UM_TIME_GIORNI, UM_TIME_CENTIORE, UM_TIME_CENTESIMI, UM_TIME_CADENZA, UM_TIME_SECONDI
        vntTempo = Str(vntTempo)
    End Select
    strOut$ = strOut$ & vntTempo & parTpS _
        & cTime.UM2Dsc(enmUM) & parTpS
Return

End Function

Sub RiordinaCollectionDBA1(ArtComposto As MXBusiness.CComposto, colListaComponenti As Collection)
Dim ArtComponente As MXBusiness.CComponente

    For Each ArtComponente In ArtComposto.pComponentiPropri
        'inserimento artcomponente su collection
        colListaComponenti.Add ArtComponente
        If (ArtComponente.pDatiComposto.pComponentiPropri.Count > 0) Then
            Call RiordinaCollectionDBA1(ArtComponente.pDatiComposto, colListaComponenti)
        End If
    Next ArtComponente
    
End Sub

Function KeyControlloGet(ctrlGen As Control) As String
Dim vntIndex As Variant
    On Local Error Resume Next
    vntIndex = ctrlGen.Index
    If (Err = 0) Then
        KeyControlloGet = ctrlGen.Name & "_" & vntIndex
    Else
        KeyControlloGet = ctrlGen.Name
    End If
    On Local Error GoTo 0
End Function

' rif.sch A4888
' Condiviso il sorgente con METCOMSTD per l'esecuzione di tutti gli agenti
#If BATCH <> 1 Then
'NOME           : FormImpostaAgenti
'DESCRIZIONE    : imposta gli agenti da eseguire
'PARAMETRO 1    : form in cui impostare gli agenti
Public Function FormImpostaAgenti(frmMyDef As Form) As Long

    If Not (frmMyDef Is Nothing) Then
        'impostazione agenti
        Set frmDefAge.frmDef = frmMyDef
        frmDefAge.Show vbModal
    End If

End Function
'NOME           : MostraRIFControlli
'DESCRIZIONE    : mostra riferimenti ai controlli
Public Sub MostraRIFControlli(frm As Form)
Dim ctrlGen As Control

    frm.Enabled = True
    On Local Error Resume Next
    For Each ctrlGen In ControlliForm(frm)
        ctrlGen.ToolTipText = RIFControlloGet(ctrlGen)
    Next
    Set ctrlGen = Nothing
    On Local Error GoTo 0
    
End Sub


Private Function RIFControlloGet(ctrlGen As Control) As String
    Dim strRif As String

    strRif = " #" & KeyControlloGet(ctrlGen) & "  "
    RIFControlloGet = strRif
    
End Function

#End If


