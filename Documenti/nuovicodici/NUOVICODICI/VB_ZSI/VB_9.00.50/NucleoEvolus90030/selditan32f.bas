Attribute VB_Name = "MSelDitAnBaf"
Option Explicit
DefInt A-Z
Private mBolCambioUtente As Boolean
Private bolCloseSelection As Boolean
Private MBolSaltaMessaggiConnessione As Boolean  'Per evitare i messaggi di riconnessione in caso di annullamento della selezione Ditta su Evolus (segnalazione Evolus Nr. 27)

Public InSelezioneDitta As Boolean

'Rif anomalia #3160
Global bolTrustedConnection As Boolean

#If ISM98SERVER <> 1 Then

    Public Function CambioUtenteAttivo() As Boolean
        Dim strUtente As String
        Dim strPassword As String
        
        mBolCambioUtente = True
        #If ISMETODO2005 <> 1 Then
            'scarico la form moduli affinche' il menu di rapido di metodo xp si possa salvare correttamente
            Unload frmModuli
        #End If
        'CRISTIAN: Cancellazione del nomemacchina dalla TabUtenti in cambiamento dell'utente attivo
        Select Case UCase(MXNU.EXEName)
            Case "METODO98", "METODOXP", "METODOEVOLUS"
                Call MXDB.dbAggiornaTabUtenti(hndDBArchivi, False)
        End Select
        
        #If ISMETODO2005 = 1 Then   'Spostato prima di ApriLoginUtente per Anomalia 10232
            Call frmModuli.SalvaPreferitiUtente
            Call frmShortcuts.SalvaShortcuts
            '<rif anomalia 8930 RZ>
            Set frmLog = Nothing
            Call metodo.SalvaLayout
        #Else
            'Anomalia 7868
            metodo.Barra.Buttons(idxBottoneModuli).Enabled = False
        #End If
        
        If (MXDB.ApriLoginUtente(strUtente, strPassword)) Then
            MbolInChiusura = False
            
            Call ChiudiDitta
            
            If Not ApriDittaAnno(True, strUtente, strPassword) Then
                MbolInChiusura = True
                'Rif. Anomalia Nr. 7688
                GBolNoMsgConfermaUscita = True
                Unload metodo
                Exit Function
            Else
                metodo.MousePointer = vbHourglass
                Call CaricaModuliMetodo(True)
                #If ISMETODO2005 = 1 Then
                    Call frmShortcuts.CaricaShortcuts
                    Call metodo.CaricaLayout
                    'Anomalia 8948
                    DoEvents
                    ' S#3040 - rimossa la Gestione Accessi da Evolus
                    'metodo.Barra.Buttons(idxBottoneDefAccessi).Enabled = (Not MXNU.CtrlAccessi)
                #Else
                    'Anomalia 7868
                    metodo.Barra.Buttons(idxBottoneModuli).Enabled = True
                #End If
                frmModuli.ModuloAttivo = "Metodo98"
                Call AggiornaStatusBar
                metodo.MousePointer = vbNormal
            End If
            CambioUtenteAttivo = True
        Else
            'Call ChiudiMetodo
            'End
            'Rif. Anomalia Nr. 7688
            GBolNoMsgConfermaUscita = True
            Unload metodo
            Exit Function
        End If
        DoEvents
        mBolCambioUtente = False
        
    End Function


    Public Function ApriDittaAnno(bolAperturaMetodo As Boolean, strUtente As String, strPassword As String) As Boolean
        Dim strDitta As String
        Dim bolOk As Boolean
        Dim bolVersione As Boolean
        
        If CmbDittaBusy Then
            Call MXNU.MsgBoxEX(3178, vbExclamation, 1007)
            ApriDittaAnno = False
            Exit Function
        End If
        
        If Not bolAperturaMetodo Then
            If Not MXNU.ISMETODO2005 Then
                frmModuli.Enabled = False
                metodo.Enabled = False
                If SelezioneDitta(strDitta) Then
                    Call ChiudiDitta
                    MXNU.DittaAttiva = strDitta
                    ' Anomalia n.ro 5242 e n.ro 5329
                    #If ISNUCLEO = 0 Then
                        If Not (Designer Is Nothing) Then
                            bolVersione = Designer.AttivaVersione
                        End If
                    #End If
                Else
                    ApriDittaAnno = False
                    frmModuli.Enabled = True
                    metodo.Enabled = True
                    Exit Function
                End If
            Else
                #If ISMETODO2005 = 1 Then
                    Call frmModuli.SalvaPreferitiUtente
                    Call frmShortcuts.SalvaShortcuts
                    '<rif anomalia #8930 RZ>
                    Set frmLog = Nothing
                    Call metodo.SalvaLayout
                #End If
                ' Rif. anomalia #3160 per Metodo Evolus
                If MXNU.LoginIntegrato Then
                    If SelezioneDitta(strDitta) Then
                        Call ChiudiDitta
                        MXNU.DittaAttiva = strDitta
                        ' Anomalia n.ro 5242 e n.ro 5329
                        #If ISNUCLEO = 0 Then
                            If Not (Designer Is Nothing) Then
                                bolVersione = Designer.AttivaVersione
                            End If
                        #End If
                    Else
                        Call ChiudiDitta
                        ApriDittaAnno = False
                        frmModuli.Enabled = True
                        metodo.Enabled = True
                        Exit Function
                    End If
                Else
                    Call ChiudiDitta
                End If
                ' Fine rif. anomalia #3160 per Metodo Evolus
            End If
        End If
        
        bolOk = True
        Do
            If ApriDitta(strUtente, strPassword) Then
                Dim NuovoAnno As Integer
                If SelezioneAnno(False, NuovoAnno) Then
                    MXNU.AnnoAttivo = NuovoAnno
                    Call ApriAnno(Not bolAperturaMetodo)
                    bolOk = True

                    #If USAM98SERVER Then
                    If MXNU.LeggiProfilo(MXNU.DirAvvio & "\mw.ini", "METODOW", "SERVIZIOMET", 0) = 0 Then
                        Call SottoponiSessione(GobjM98Server)
                    Else
                        Set GobjM98Server = Nothing
                        Call ConnectM98Server(GobjM98Server)
                    End If
                    #End If
                    
                    If Not bolAperturaMetodo Then
                        On Local Error Resume Next
                        Call MXNU.RicaricaRisorseDitta ' Rif scheda n.ro 3092 (Rif. Scheda n.ro 1 Anomalia ISV)
                        frmModuli.ModuloAttivo = "Metodo98"
                        If Not MXNU.ISMETODO2005 Then
                            Unload frmModuli
                            Call CaricaModuliMetodo
                        Else
                            Call CaricaModuliMetodo(True)
                            #If ISMETODO2005 = 1 Then
                                Call metodo.CaricaLayout
                                
                            #End If
                        End If
                        frmModuli.ModuloAttivo = "Metodo98"
                        On Local Error GoTo 0
                    End If
                Else
                    bolOk = False
                End If
            Else
                bolOk = False
            End If
            If Not bolOk Then
                If MXNU.MsgBoxEX(9011, vbCritical + vbYesNo, 1007) = vbNo Then
                    ApriDittaAnno = False
                    bolCloseSelection = True
                    Exit Do
                Else
                    Call ChiudiDitta
                    MXNU.DittaAttiva = ""
                    ApriDittaAnno = ApriDittaAnno(bolAperturaMetodo, MXNU.UtenteAttivo, MXNU.PasswordUtente)
                    'sono riuscito ad aprire la ditta
                    If ApriDittaAnno Or bolCloseSelection Then Exit Do
                End If
            Else
                ApriDittaAnno = True
                ' Anomalia n.ro 5242 e n.ro 5329
#If ISNUCLEO = 0 Then
                If Not (Designer Is Nothing) Then
                    bolVersione = Designer.AttivaVersione
                End If
#End If
                Exit Do
            End If
        Loop
                        
        ' Rif. anomalia n.ro 4824
        #If ISMETODO2005 = 1 Then
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneDefAgenti).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneNomiCtrlCmp).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneSituazAnagr).Enabled = Not (MXNU.CtrlAccessi)
        #Else
            With metodo.Barra.Buttons
                .Item(idxBottoneDefAgenti).ButtonMenus.Item(1).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
                .Item(idxBottoneDefAgenti).ButtonMenus.Item(2).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
                .Item(idxBottoneDefAgenti).ButtonMenus.Item(3).Enabled = Not (MXNU.CtrlAccessi)
            End With
        #End If
        ' Fine rif. anomalia n.ro 4824

        If Not bolAperturaMetodo Then
            frmModuli.Enabled = True
            metodo.Enabled = True
        End If
        
    End Function


    Public Function SelezioneAnno(bolSel As Boolean, intNuovoAnno As Integer) As Boolean
        Dim colRisultati As Collection, intq As Integer, bolEseguiSel As Boolean
        SelezioneAnno = True
        If bolSel Then
            bolEseguiSel = True
        Else
            bolEseguiSel = (MXNU.AnnoAttivo = 0) Or Not EsisteAnno()
        End If
        If bolEseguiSel Then
            If MXNU.FrmMetodo Is Nothing Then Set MXNU.FrmMetodo = frmIntro
            If (MXVI.Selezione("TABESE", "CODICE", "", False, Nothing, colRisultati)) Then
                'MXNU.AnnoAttivo = colRisultati(1)("Codice")  '<=== Remmato per sk anomalie 5113: AnnoAttivo deve essere modificato DOPO ChiudiFormAttive
                intNuovoAnno = colRisultati(1)("Codice")
            Else
                intNuovoAnno = MXNU.AnnoAttivo
                SelezioneAnno = False
            End If
        Else
            intNuovoAnno = MXNU.AnnoAttivo
            intq = MXDB.dbEseguiSQL(hndDBArchivi, "UPDATE TabUtenti SET EsercizioAttivo=" & MXNU.AnnoAttivo & " WHERE UserID='" & MXNU.UtenteAttivo & "' AND NOMEMACCHINA=" & hndDBArchivi.FormatoSQL(MXNU.NomeComputer, DB_TEXT))
        End If
        Set colRisultati = Nothing
        
    End Function
    
    Public Function SelezioneDitta(strDitta As String) As Boolean
    
        SelezioneDitta = False
        If (frmSelDitta.SelezioneDitta(strDitta)) Then
            SelezioneDitta = True
        End If
    
    End Function

    
    Sub CancellaDitta()
        Dim strDitta As String
        
        On Local Error GoTo CancellaDitta_Err
        If SelezioneDitta(strDitta) Then
            If strDitta <> MXNU.DittaAttiva Then
                If MXNU.MsgBoxEX(1860, vbQuestion + vbYesNo, 1007, strDitta) = vbYes Then
                     MXNU.ScriviProfilo MXNU.PercorsoLocal$ & "\ditte.ini", "DITTE", strDitta, 0&
                     MXNU.ScriviProfilo MXNU.PercorsoLocal$ & "\ditte.ini", "CONNESSIONE", strDitta, 0&
                     Call MXNU.MsgBoxEX(2368, vbInformation, 1007)
                End If
            End If
        End If
CancellaDitta_Fine:
        On Local Error GoTo 0
        Exit Sub
        
CancellaDitta_Err:
        Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("CancellaDitta", Err.Number, Err.Description))
        Resume CancellaDitta_Fine
    End Sub
    

#End If

Public Sub ApriAnno(bolAggSts As Boolean)
    Dim intq As Integer
    Dim strSQL As String
    Dim HSS As CRecordSet

    strSQL = "SELECT * FROM TabEsercizi WHERE Codice=" & MXNU.AnnoAttivo
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    
    MXNU.DescrizioneAnnoAttivo = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "Descrizione", 0)
    MXNU.DataIniCont = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "DataIniCont", 0)
    MXNU.DataFineCont = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "DataFineCont", 0)
    MXNU.DataIniMag = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "DataIniMag", 0)
    MXNU.DataFineMag = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "DataFineMag", 0)
    MXNU.DataIniIva = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "DataIniIva", 0)
    MXNU.DataFineIva = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "DataFineIva", 0)
    MXNU.UsaEuro = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "UsaEuro", 0)
    MXNU.StatoEsercizioCont = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "StatoCont", 0)
    MXNU.StatoEsercizioMag = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "StatoMag", 0)
    MXNU.LiqIva = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "LiqIva", 0)
    MXNU.IntraStatAcq = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "IntraStatAcq", 0)
    MXNU.IntraStatVend = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "IntraStatVend", 0)
    MXNU.IntraRegimeAcq = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "IntraRegimeAcq", 0)
    MXNU.IntraRegimeVend = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "IntraRegimeVend", 0)
    
    intq = MXDB.dbChiudiSS(HSS)

    #If ISM98SERVER <> 1 Then
        intq = MXDB.dbEseguiSQL(hndDBArchivi, "UPDATE TabUtenti SET EsercizioAttivo=" & MXNU.AnnoAttivo & " WHERE UserID='" & MXNU.UtenteAttivo & "' AND NOMEMACCHINA=" & hndDBArchivi.FormatoSQL(MXNU.NomeComputer, DB_TEXT))
    #Else
        intq = MXDB.dbEseguiSQL(hndDBArchivi, "UPDATE TabUtenti SET EsercizioAttivo=" & MXNU.AnnoAttivo & " WHERE UserID='" & MXNU.UtenteAttivo & "'")
    #End If
    
    Call MXDB.DBSetConnProperty(hndDBArchivi.ConnessioneW)
    
    #If ISKEY <> 1 Then
        Call LeggiVincoli
        Call LeggiVincoliMagazzino
    #End If
    'resetto i vincoli di produzione (rif.sch.1830)
    MXDBA.ResettaVincoliDisinta
    MXCICLI.ResettaVincoliCiclo
    MXRIS.ResettaVincoliRisorse
    MXPROD.ResettaVincoliProduzione
    
    #If ISM98SERVER <> 1 Then
        If (bolAggSts) Then Call AggiornaStatusBar
        Call MXNU.SalvaImpostazioniUtente(MXNU.UtenteSistema)
    #End If

End Sub

Sub ChiudiDitta()
    Dim bAggUtenti As Boolean
'    If Not CmbDittaBusy Then
        #If ISM98SERVER <> 1 Then
            Call ChiudiFormAttive
        #End If
        If (Left(UCase(MXNU.EXEName), 6) = "METODO") And Not mBolCambioUtente Then
            bAggUtenti = False
            If Not (hndDBArchivi Is Nothing) Then
                bAggUtenti = (hndDBArchivi.ConnessioneR.State = adStateOpen)
            End If
            If bAggUtenti Then Call MXDB.dbAggiornaTabUtenti(hndDBArchivi, False)
        End If
        Call MXVA.ChiudiDyTRAnagraf
        Call MXVA.ChiudiDyTRValidazione
        Call MXCT.ChiudiDyTRTabelle
        Call MXVI.ChiudiDyTRVisioni
        Call MXVI.ChiudiDyTRSituazioni
        #If ISM98SERVER <> 1 Then
             If bAggUtenti Then Call MXNU.SalvaImpostazioniUtente(MXNU.UtenteSistema)
        #End If
 '   Else
  '      MsgBox "Cambio ditta in corso, chiudere la form"
  '  End If
End Sub

Function ApriDitta(strUtente As String, strPassword As String) As Boolean
    Dim strDitta As String
    Dim intTentativi As Integer
    Dim bolConnesso As Boolean
    Dim HSS As MXKit.CRecordSet
    Dim lngVersione As Long
    Dim intq As Integer
    Dim strLog As String
    Dim strDesErr As String
    Dim strLineErr As String
    On Local Error GoTo ApriDitta_Err

    ApriDitta = True
RitentaConnessione:
    'rif. anomalia #3160
    #If ISM98SERVER <> 1 Then
        If strUtente <> "" Then
            Set hndDBArchivi = MXDB.dbApriDB(NOME_DB_ARCHIVI, strUtente, strPassword, hndDBArchivi)
        Else
            Set hndDBArchivi = MXDB.dbApriDB(NOME_DB_ARCHIVI, , , hndDBArchivi)
        End If
        bolConnesso = (Not (hndDBArchivi Is Nothing))
    #Else
        If strUtente <> "" And Not (bolTrustedConnection) Then ' <-- per rif. anomalia #3160 aggiunto variabile booleana di controllo
            Set hndDBArchivi = MXDB.dbApriDB(NOME_DB_ARCHIVI, strUtente, strPassword, hndDBArchivi)
        Else
            Set hndDBArchivi = MXDB.dbApriDB(NOME_DB_ARCHIVI, , , hndDBArchivi)
        End If
        bolConnesso = (Not (hndDBArchivi Is Nothing))
    #End If
    
    If bolConnesso Then bolConnesso = (hndDBArchivi.ConnessioneR.State <> adStateClosed)
    #If ISM98SERVER <> 1 Then
        
        If Not bolConnesso Then
            'If (MXNU.MsgBoxEX(9011, vbYesNo + vbQuestion + vbDefaultButton1, 1007) = vbYes) Then
            'Utilizzo MsgBox standard di vb perchè in alcuni casi non visualizza il msgbox e risponde automaticamente no
            Dim r As VbMsgBoxResult
            #If ISMETODO2005 = 1 Then
                If InSelezioneDitta Then
                    r = vbNo
                Else
                    r = MsgBox(MXNU.CaricaStringaRes(9011), vbYesNo + vbQuestion + vbDefaultButton1, MXNU.CaricaStringaRes(1007))
                End If
            #Else
                r = MsgBox(MXNU.CaricaStringaRes(9011), vbYesNo + vbQuestion + vbDefaultButton1, MXNU.CaricaStringaRes(1007))
            #End If
            If r = vbYes Then
                If MXNU.ISMETODO2005 Then
                    GoTo RitentaConnessione
                Else
                    If (SelezioneDitta(strDitta)) Then
                        MXNU.DittaAttiva = strDitta
                        Call MXNU.SalvaImpostazioniUtente(MXNU.UtenteSistema)
                        GoTo RitentaConnessione
                    Else
                        If MXNU.FrmMetodo Is Nothing Then
                            Call ChiudiMetodo
                        Else
                            GoTo RitentaConnessione
                        End If
                        ApriDitta = False
                        Exit Function
                    End If
                End If
            Else
                If MXNU.FrmMetodo Is Nothing Then
                    Call ChiudiMetodo
                Else
                    If Not MXNU.ISMETODO2005 Then
                        GoTo RitentaConnessione
                    End If
                End If
                ApriDitta = False
                Exit Function
            End If
        End If
        '>>> LOGIN UTENTE
        If MXNU.UtenteDB = "" Then
            ApriDitta = False
            GoTo ApriDitta_Fine
        End If
        'Controllo Versione
        Call MXDB.dbClearUltimoErrore
        strLineErr = "Lettura TabVersioni"
        strLog = MXNU.GetTempFile()
        Call MXNU.ImpostaErroriSuLog(strLog, True)
        Set HSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT Max(Codice) MaxCodice FROM TABVERSIONIM98")
        If MXDB.dbUltimoErrore(strDesErr) = 0 Then
            If Not MXDB.dbFineTab(HSS) Then
                lngVersione = MXDB.dbGetCampo(HSS, NO_REPOSITION, "MaxCodice", 0)
            End If
            intq = MXDB.dbChiudiSS(HSS)
            Call MXNU.ChiudiErroriSuLog
            If Val(Replace(MXNU.VersioneMetodo, ".", "")) <> lngVersione Then
                If Not MBolSaltaMessaggiConnessione Then  'Per MetodoEvolus, nel caso si faccia annulla della selezione ditte, viene rifatta la connessione ad db precedente
                    Call MXNU.MsgBoxEX(1186, vbCritical, 1007, Array("", Val(Replace(MXNU.VersioneMetodo, ".", "")), lngVersione))
                End If
            End If
        Else
            intq = MXDB.dbChiudiSS(HSS)
            Call MXNU.ChiudiErroriSuLog
            Call MXNU.MsgBoxEX(1185, vbCritical, 1007, Array("", strDesErr))
        End If
        strLineErr = "Copia File mwpers.ini"
        '>>> FILE INI PERSONALE
        If Dir$(MXNU.File_ini_personale, vbNormal) = "" Then
            FileCopy MXNU.PercorsoPreferenze & "\mwpers.ini", MXNU.File_ini_personale
        End If
        strLineErr = "Copia File mwpersvis.ini"
        If Dir$(MXNU.File_ini_personaleVisioni, vbNormal) = "" Then
            FileCopy MXNU.PercorsoPreferenze & "\mwpersvis.ini", MXNU.File_ini_personaleVisioni
            Call SpostaSezioneVisioni
        End If
        strLineErr = "Salva Impostazioni Utente"
        Call MXNU.SalvaImpostazioniUtente(MXNU.UtenteSistema)

        MXNU.Dsc_Breve_Ditta = LeggiDscBreve(MXNU.DittaAttiva)
        #If TOOLS <> 1 And ISNUCLEO = 0 Then
        'Sviluppo nr. 773
        strLineErr = "Lettura Record ComandiBatch"
        If EsistonoComandiBatch("DATAORA<{ fn Now()}-1") Then
            If Not MBolSaltaMessaggiConnessione Then  'Per MetodoEvolus, nel caso si faccia annulla della selezione ditte, viene rifatta la connessione ad db precedente
                Call MXNU.MsgBoxEX(2585, vbCritical, 1007)
            End If
        End If
        #End If
    #Else
        If MXNU.UtenteDB = "" Then
            ApriDitta = False
            GoTo ApriDitta_Fine
        End If
        ApriDitta = bolConnesso
        If bolConnesso Then
            Dim hssDesDitta As MXKit.CRecordSet
            Dim strSQL As String
            Dim strDes As String
            
            strSQL = "SELECT DataCostituzione,DesBreve FROM  TabDitte"
            Set hssDesDitta = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
            If Not MXDB.dbFineTab(hssDesDitta, TIPO_SNAPSHOT) Then
                strDes = MXDB.dbGetCampo(hssDesDitta, TIPO_SNAPSHOT, "DesBreve", "")
                MXNU.DataCostituzione = MXDB.dbGetCampo(hssDesDitta, TIPO_SNAPSHOT, "DataCostituzione", MXNU.Default_Data)
            Else
                strDes = ""
            End If
            Call MXDB.dbChiudiSS(hssDesDitta)
            MXNU.Dsc_Breve_Ditta = strDes
        Else
            GoTo ApriDitta_Fine
        End If
    #End If
    
    MXNU.DSNDittaAttiva = MXDB.UltimoDSNAperto
    
    strLineErr = "Lettura licenze software"
    '[15/06/2011] Rimozione Chiave Hardware - Controllo spostato all'interno di ApriDitta
    If (MXNU.ChiaveSoftwarePresente(MXNU.DittaAttiva, hndDBArchivi.ConnessioneR) <> EnmStatoChiaveSoftware.ChiavePresente) Then
        Call MXNU.MsgBoxEX(9000, vbOKOnly + vbCritical, 1007)
        ApriDitta = False
        GoTo ApriDitta_Fine
    End If
    
    'ricarico tutti i dynaset temporanei dei file .DAT
    strLineErr = "Caricamento Validazioni"
    Call MXVA.ApriDyTRValidazione
    strLineErr = "Caricamento Anagrafiche"
    Call MXVA.ApriDyTRAnagraf
    strLineErr = "Caricamento Tabelle"
    Call MXCT.ApriDyTRTabelle
    strLineErr = "Caricamento Visioni"
    Call MXVI.ApriDyTRVisioni
    strLineErr = "Caricamento Situazioni"
    Call MXVI.ApriDyTRSituazioni
    
ApriDitta_Fine:
    On Local Error GoTo 0
    Exit Function
    
ApriDitta_Err:
    Call MXNU.MsgBoxEX(9010, vbCritical, 1007, Array("ApriDitta", Err.Number, Err.Description & " [" & strLineErr & "]"))
    On Local Error GoTo 0
    ApriDitta = False
    Call ChiudiDitta
    #If ISM98SERVER <> 1 Then
        If MXNU.FrmMetodo Is Nothing Then
            Call ChiudiMetodo
        Else
            Unload MXNU.FrmMetodo
        End If
    #End If
    Resume ApriDitta_Fine
End Function


Function EsisteAnno() As Boolean
    Dim HSS As CRecordSet, intq As Integer, strSQL As String
    
    EsisteAnno = False
    If MXNU.AnnoAttivo > 0 Then
        strSQL = "SELECT * FROM TabEsercizi WHERE Codice=" & MXNU.AnnoAttivo
        Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
        EsisteAnno = Not MXDB.dbFineTab(HSS)
        intq = MXDB.dbChiudiSS(HSS)
    End If
    
End Function




Sub LeggiVincoli()
Dim q As Integer
Dim hndtn As CRecordSet
Dim inti As Integer
Dim strSQL As String
Dim hTabEse As MXKit.CRecordSet     'Tabella esercizi


    Set hndtn = MXDB.dbCreaSS(hndDBArchivi, "SELECT * FROM TabVincoliGIC WHERE Esercizio=" & MXNU.AnnoAttivo, TIPO_TABELLA)
    
    For inti = 1 To 5
        MXNU.VincoliIva(IVA_VEN, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVADeb" & CStr(inti), "")
        MXNU.VincoliIva(IVA_ACQ, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVACred" & CStr(inti), "")
        MXNU.VincoliIva(IVA_SOS, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVASosp" & CStr(inti), "")
        MXNU.VincoliIva(IVA_VENINTRA, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVAVendIntra" & CStr(inti), "")
        MXNU.VincoliIva(IVA_ACQINTRA, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVAAcqIntra" & CStr(inti), "")
        MXNU.VincoliIva(IVA_AUTOFATTURE, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVAAutoFatt" & CStr(inti), "")   'Sviluppo 1368
        MXNU.VincoliIva(IVA_SOS_CRED, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIvaSospCred" & CStr(inti), "")
    Next inti
    
    MXNU.Vincoli(SC_CLI_CORRISP) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCCliCorrisp", "")
    MXNU.Vincoli(REG_IVA_INTRA) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "RegVendIntra", ""))
    MXNU.Vincoli(CAUS_INSOLUTO) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "CausContInsoluto", ""))
    MXNU.Vincoli(CAUS_APERTURA) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "CausContAp", ""))
    MXNU.Vincoli(CAUS_CHIUSURA) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "CausContCh", ""))
    MXNU.Vincoli(CONTO_PATR_APERTURA) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoPatrAP", "")
    MXNU.Vincoli(CONTO_PATR_CHIUSURA) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoPatrCH", "")
    MXNU.Vincoli(CONTO_ECO_CHIUSURA) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoEcoCH", "")
    MXNU.Vincoli(REG_IVA_AUTOFATT) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "RegVendAutoFatt", ""))   'Sviluppo 1368
    
    'Rif. Sviluppo 589
    MXNU.Vincoli(CAUS_UTILEPERDITA) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "CausContRilUtPerd", ""))   'Sviluppo 1368
    MXNU.Vincoli(CONTO_UTILE_ESERCIZIO) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoUtileEserc", ""))   'Sviluppo 1368
    MXNU.Vincoli(CONTO_PERDITA_ESERCIZIO) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoPerdEserc", ""))   'Sviluppo 1368
    
    MXNU.CodCambioLire = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "DivisaLire", 0)
    MXNU.CodCambioEuro = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "DivisaEuro", 0)
    MXNU.DecimaliQuantita = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "nDecimaliQuantita", 0)
    MXNU.DecimaliPesiVolumi = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "nDecimaliPesiVol", 0)
    MXNU.DecimaliLireTotale = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "nDecimaliTotaleLire", 0)
    MXNU.DecimaliLireUnitario = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "nDecimaliUnitarioLire", 0)
    MXNU.FORMATO_QUANTITA = Formato("####,###,##0", MXNU.DecimaliQuantita)
    MXNU.FORMATO_PESIVOLUMI = Formato("####,###,##0", MXNU.DecimaliPesiVolumi)
    MXNU.FORMATO_LIRE_UNITARIO = Formato("####,###,###,##0", MXNU.DecimaliLireUnitario)
    MXNU.FORMATO_LIRE_TOTALE = Formato("####,###,###,##0", MXNU.DecimaliLireTotale)
    'Sviluppo nr. 1566
    MXNU.ImportiSpRipMag = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "IncludiSpRip", 0)
    
    'Imposto nella proprietà del nucleo l'ultimo anno creato.
    Set hTabEse = MXDB.dbCreaSS(hndDBArchivi, "SELECT MAX(CODICE) AS ULTESE FROM TABESERCIZI")
    MXNU.UltimoEsercizioCreato = MXDB.dbGetCampo(hTabEse, TIPO_SNAPSHOT, "ULTESE", MXNU.AnnoAttivo)
    Call MXDB.dbChiudiSS(hTabEse)
    
    q = MXDB.dbChiudiSS(hndtn)
    
    If MXNU.CodCambioLire = MXNU.CodCambioEuro Then
        Call MXNU.MsgBoxEX(1399, vbCritical, 1007)
    End If
    Call GetFormatiEuro
    
    'lettura vincoli produzione
    strSQL = "select NDECIMALICICLO from TABVINCOLIPRODUZIONE order by PROGRESSIVO desc"
    Set hndtn = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    'RIF.A#6402 - memorizzo il numero di decimali dei centesimi
    MXNU.DecimaliCentesimi = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "NDECIMALICICLO", 0)
    MXNU.Formato_Centesimi = Formato("####,###,##0", MXNU.DecimaliCentesimi)
    q = MXDB.dbChiudiSS(hndtn)
End Sub


Sub GetFormatiEuro()
    Dim HSS As CRecordSet, intDec As Integer, intq As Integer
    
    MXNU.FORMATO_EURO_UNITARIO = "###,###,##0.00"
    MXNU.FORMATO_EURO_TOTALE = "###,###,##0.00"
    MXNU.DecimaliEuroTotale = 2
    MXNU.DecimaliEuroUnitario = 2
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT nDecimaliUnitario,nDecimaliTotale FROM TabCambi WHERE Codice=(SELECT DivisaEuro FROM TabVincoliGIC WHERE Esercizio=" & MXNU.AnnoAttivo & ")", TIPO_TABELLA)
    If Not MXDB.dbFineTab(HSS, TIPO_DYNASET) Then
        intDec = MXDB.dbGetCampo(HSS, NO_REPOSITION, "nDecimaliUnitario", 0)
        MXNU.DecimaliEuroUnitario = intDec
        MXNU.FORMATO_EURO_UNITARIO = Formato("####,###,###,##0", intDec)
        
        intDec = MXDB.dbGetCampo(HSS, NO_REPOSITION, "nDecimaliTotale", 0)
        MXNU.FORMATO_EURO_TOTALE = Formato("####,###,###,##0", intDec)
        MXNU.DecimaliEuroTotale = intDec
    End If
    intq = MXDB.dbChiudiSS(HSS)
    
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT CambioEuro FROM TabCambi WHERE Codice=0")
    If Not MXDB.dbFineTab(HSS, TIPO_SNAPSHOT) Then
        MXNU.CambioLireEuro = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "CambioEuro", 1)
    Else
        MXNU.CambioLireEuro = 1
    End If
    intq = MXDB.dbChiudiSS(HSS)
    
    
End Sub

Function Formato(strDes As String, intDec As Integer) As String
    Dim strD As String
    
    If intDec > 0 Then
        strD = Right(strDes & Left(".000000", intDec + 1), Len(strDes))
    Else
        strD = strDes
    End If
    If Left(strD, 1) = "," Then
        Mid(strD, 1, 1) = "#"
    End If
    Formato = strD

End Function


Private Sub SpostaSezioneVisioni()
    Dim strRiga As String
    Dim vetStr() As String
    Dim lngNum As Long
    Dim i As Integer
    Dim intq As Integer
    
    strRiga = MXNU.LeggiProfilo(MXNU.File_ini_personale, "VISIONI", 0&, "")
    If strRiga <> "" Then
        vetStr = Split(strRiga, vbNullChar, , vbTextCompare)
        lngNum = UBound(vetStr)
        For i = 0 To lngNum
            strRiga = MXNU.LeggiProfilo(MXNU.File_ini_personale, "VISIONI", vetStr(i), "")
            Call MXNU.ScriviProfilo(MXNU.File_ini_personaleVisioni, "VISIONI", vetStr(i), strRiga)
        Next i
        Call MXNU.ScriviProfilo(MXNU.File_ini_personale, "VISIONI", 0&, "")
    End If
    
End Sub
Public Sub ControllaOpSchedulate()

#If BATCH <> 1 And ISM98SERVER <> 1 And TOOLS <> 1 And ISNUCLEO = 0 Then
    Dim oSched As MxScheduler.clsScheduler
    Set oSched = New MxScheduler.clsScheduler
    If oSched.Inizializza(MXNU, Command()) Then
        If oSched.SegnalaLog(MXNU) Then
            Call oSched.MostraLogNonLetti(MXNU, False)
        End If
    End If
#End If

'    Dim oRs As ADODB.RecordSet
'    Dim strSQL As String
'    Dim log As CGestLog
'    Dim oSched As MxScheduler.clsScheduler
'    Dim strLog As String, strLogFile As String
'    Dim strDescrOp As String
'    Dim strMsg As String
'
'    Set oSched = New MxScheduler.clsScheduler
'    If oSched.Inizializza(MXNU, Command()) Then
'        If oSched.SegnalaLog(MXNU) Then
'            Set log = New CGestLog
'            strLog = MXNU.GetTempFile()
'            Call MXNU.ImpostaErroriSuLog(strLog, True)
'
'            strSQL = "SELECT logop.idop, logop.dataop, logop.ditta, logop.logfile, logop.letto, oper.descrop, oper.nomeop" & vbCrLf & _
'                "  FROM LOGOPERAZIONI logop, OPERAZIONI oper" & vbCrLf & _
'                "  Where logop.idop = oper.idop AND logop.letto=0 AND oper.utente=" & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT)
'            Set oRs = New ADODB.RecordSet
'            Call oRs.Open(strSQL, oSched.mobjScheduler.GetAdo.GetConnection, adOpenForwardOnly, adLockReadOnly)
'
'            Do While Not oRs.EOF
'                strDescrOp = oRs("descrop").value
'                strLogFile = oRs("logfile").value
'                strMsg = strDescrOp & ": " & strLogFile
'                Call MXNU.MsgBoxEX(strMsg, vbInformation, "JobScheduler")
'                oRs.MoveNext
'            Loop
'            oRs.Close
'            Set oRs = Nothing
'
'            Call MXNU.ChiudiErroriSuLog
'            Call log.MostraFileLog(strLog)
'            'Call oSched.MostraLogNonLetti(MXNU)
'        End If
'    End If
'    Set oSched = Nothing
End Sub

Public Sub ApriSchedulatore()

#If BATCH <> 1 And ISM98SERVER <> 1 And TOOLS <> 1 And ISNUCLEO = 0 Then
    Dim oSched As MxScheduler.clsScheduler
    Set oSched = New MxScheduler.clsScheduler
    If oSched.Inizializza(MXNU, Command()) Then
        Call oSched.GestisciOperazioni(False)
    End If
    
#End If
End Sub


Public Sub OldControllaOpSchedulate()
#If BATCH <> 1 And ISM98SERVER <> 1 And TOOLS <> 1 And ISNUCLEO = 0 Then
    Dim strFile As String
    Dim q As Integer
    Dim strTitolo As String
    
    On Local Error GoTo OP_Err
    
    strFile = Dir$(MXNU.PercorsoPreferenze & "\SCHEDULA\" & CStr(MXNU.NTerminale) & "_*.log")
    While strFile <> ""
        q = InStr(strFile, CStr(MXNU.NTerminale) & "_")
        strTitolo = Mid$(strFile, q + 2)
        q = InStr(strTitolo, ".")
        strTitolo = Left$(strTitolo, q - 1)
        strTitolo = MXNU.CaricaCaptionInLingua(MXNU.LeggiProfilo(MXNU.PercorsoPgm & "\MWSchedula.ini", "DESCRIZIONE", strTitolo, strTitolo))
        #If ISMETODO2005 = 1 Then
            frmLog.MstrTitolo = strTitolo
            Call frmLog.MostraFileLog(MXNU.PercorsoPreferenze & "\SCHEDULA\" & strFile, , , True)
        #Else
            Dim frmLogOpSc As frmLog
            Set frmLogOpSc = New frmLog
            frmLogOpSc.MstrTitolo = strTitolo
            Call frmLogOpSc.MostraFileLog(MXNU.PercorsoPreferenze & "\SCHEDULA\" & strFile, , , True)
            Set frmLogOpSc = Nothing
        #End If
        If MXNU.MsgBoxEX(2580, vbInformation + vbYesNo, 1007, Array(strTitolo)) = vbYes Then
            Kill MXNU.PercorsoPreferenze & "\SCHEDULA\" & strFile
        End If
        strFile = Dir
    Wend
OP_Fine:
    On Local Error GoTo 0
    Exit Sub
    
OP_Err:
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("ControllaOpSchedulate", Err.Number, Err.Description))
    Resume OP_Fine
#End If
End Sub

