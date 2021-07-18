Attribute VB_Name = "Extra"
Option Explicit
DefLng A-Z

'=================================
'   dichiarazione tipi extra
'=================================
Enum setTipoDatoExtra
    extTesto = 0
    extNumerico = 1
    extDecimal = 2
    extData = 4
    extTime = 8
    extTutti = 15 'somma di tutti i precedenti
End Enum

Private Function DammiExtra(ByVal strSezione As String, ByVal intEntry As Integer) As String
    DammiExtra = MXNU.LeggiProfilo(MXNU.PercorsoPgm & "\EXTRA.INI", strSezione, CStr(intEntry), "")
End Function

'definisce il foglio degli extra di una anagrafica
'le righe vengono caricate nello stesso ordine in cui sono presenti nel file ini
Function DefExtra(AnagraficaExtra As MXKit.Anagrafica, Foglio As FPSpreadADO.fpSpread, Sezione$) As Integer
    Dim Entry$, FileIni$, strRiga As String
    Dim picPicture As StdPicture
    
    DefExtra = False
    FileIni = CercaDirFile("extra.ini", MXNU.PercorsoPers$ & "\" & MXNU.DittaAttiva & ";" & MXNU.PercorsoPers$)
    If UCase(Dir$(FileIni$, vbNormal)) = "EXTRA.INI" Then
        Entry = MXNU.LeggiProfilo(FileIni$, Sezione$, 0&, "")
        If Entry$ <> "" Then
           Dim NumeroEntry$, ncol&, Cont%, rgtr$(), TipoCmp%, Sql$, hSS As CRecordSet, riga$, q%, lungcmp%
           Dim Row&
           
           Entry = Entry & Chr$(0)
            
           Foglio.MaxCols = 5
           Foglio.MaxRows = -1
           Call ssSpreadImposta(Foglio)
           Call ssDefStaticText(Foglio, 1, -1)
           Foglio.Col = 4
           Foglio.Row = -1
           'RIF. AN. #10180 RZ
           Call ssDefText(Foglio, 4, -1, 250)
           Foglio.Lock = True
           Call ssDefButton(Foglio, 3, -1, SS_CELL_BUTTON_NORMAL, , , ssResGetSelezione(PIC_SELEZIONE, PIC_STS_UP))
           Foglio.ColWidth(3) = 300
           Foglio.Col = 0
           Foglio.Row = -1
           Foglio.ColHidden = True
           Foglio.Col = 5
           Foglio.ColHidden = True

           On Local Error GoTo ERR_DefExtra
           While Entry <> ""
               q = InStr(Entry$, Chr$(0))
               NumeroEntry = Left(Entry, q - 1)
               Entry = Mid$(Entry, q + 1)
               riga$ = MXNU.LeggiProfilo(FileIni$, Sezione, NumeroEntry, "")
               If riga$ <> "" Then
                   Row& = Val(NumeroEntry)
                   Cont = Cont + 1
                   ReDim rgtr$(0 To 3)
                   q = slice(riga$, ",", rgtr$())
                   'Descrizione campo
                   Foglio.Row = Row&
                   Call Foglio.SetText(1, Row&, MXNU.CaricaCaptionInLingua(CStr(rgtr$(0))))
                   If rgtr(1) = "0" Then
                      Call Foglio.SetText(2, Row&, UCase$(rgtr$(2)))
                      Foglio.Col = 2
                      Foglio.Lock = True
                      On Local Error Resume Next
                      If Len(rgtr$(3)) > 0 Then
                        Set picPicture = LoadPicture(rgtr$(3))
                      End If
                      If picPicture Is Nothing Then Set picPicture = ssResGetSelezione(PIC_DETTAGLIO, PIC_STS_UP)
                      Call ssButtonSetPicture(Foglio, 3, Foglio.Row, picPicture)
                      Set picPicture = Nothing
                      On Local Error GoTo 0
                      'Call ssDefText(Foglio, 3, Foglio.Row)
                   Else
                      Foglio.Col = 3
                      Call Foglio.SetText(5, Row&, CStr(rgtr(1)))
                      If AnagraficaExtra.grinput(rgtr(1)).TipoValidazione = "" Then
                          Call ssDefText(Foglio, 3, Foglio.Row)
                          Foglio.Col = 3
                          Foglio.Lock = True
                      End If
                   End If
               End If
NEXT_DefExtra:
           Wend
           Foglio.MaxRows = Cont
           strRiga = MXNU.LeggiProfilo(MXNU.File_ini_personale, "TABELLE", "EXTRA" & Sezione$, "")
           If strRiga <> "" Then Call ssImpostazioniSet(Foglio, strRiga)
           Foglio.ReDraw = True

           DefExtra = True
       End If
    End If
    
END_DefExtra:
    On Local Error GoTo 0
    Exit Function
    
ERR_DefExtra:
    Call MXNU.MsgBoxEX(1366, vbCritical, 1007, rgtr(1))
    Resume NEXT_DefExtra
End Function

'NOME           : EsisteExtra
'DESCRIZIONE    : mi dice se esiste almeno un extra nella sezione specificata
'PARAMETRO 1    : sezione da considerare
'RITORNO        : TRUE se esistono gli extra; FALSE altrimenti
Function EsisteExtra(strSezione As String) As Boolean
    EsisteExtra = (MXNU.LeggiProfilo(MXNU.PercorsoPgm & "\EXTRA.INI", strSezione, "1", "") <> "")
End Function

'NOME           : CaricaComboExtra
'DESCRIZIONE    : carica un combo box con i campi extra
'PARAMETRO 1    : oggetto combo da inizializzare
'PARAMETRO 2    : sezione extra da leggere
'PARAMETRO 3    : tipo extra da leggere (default tutti)
'RISULTATO      : TRUE, se esistono campi extra; FALSE altrimenti
'ATTENZIONE     : Il controllo verrà settato in modo da avere la lista dei campi extra e,
'                 già impostato nell'ItemData la posizione del campo stesso.
Function CaricaComboExtra(objCombo As ComboBox, _
                            ByVal strSezione As String, _
                            Optional ByVal enmTipoExtra As setTipoDatoExtra = extTutti)
Dim strSQL As String
Dim intEntry As Integer
Dim bolFine As Boolean
Dim strExtra As String
Dim vetEntry() As String
Dim enmTipoCampo As setTipiDatiDB
Dim bolAddItem As Boolean
Dim xAna As MXKit.Anagrafica
Dim strNomeAnagrafica As String
Dim strValoriRigaExtra() As String

    On Local Error GoTo ERR_CaricaComboExtra
    intEntry = 0
    objCombo.Clear
    If (EsisteExtra(strSezione)) Then
        strSQL = QueryDefinizione(strSezione, strNomeAnagrafica)
        Set xAna = MXVA.CreaCAnagrafica(strNomeAnagrafica, Nothing)
        Do
            intEntry = intEntry + 1
            strExtra = DammiExtra(strSezione, intEntry)
            If (strExtra <> "") Then
                'Controllo che non si tratti di un extra agente (rif. anomalia 2896)
                Erase strValoriRigaExtra
                strValoriRigaExtra = Split(strExtra, ",")
                If strValoriRigaExtra(1) = "0" Then GoTo NEXT_CaricaComboExtra
            
                ReDim vetEntry(0) As String
                Call slice(strExtra, ",", vetEntry())
                
                enmTipoCampo = xAna.grinput(vetEntry(1)).TipoCampo
                bolAddItem = False
                If (enmTipoExtra = extTutti) Then
                    bolAddItem = True
                Else
                    If ((enmTipoExtra And extTesto) <> 0) Then
                        bolAddItem = bolAddItem Or (enmTipoCampo = DB_LONGVARCHAR Or enmTipoCampo = DB_TEXT)
                    End If
                    If ((enmTipoExtra And extNumerico) <> 0) Then
                        bolAddItem = bolAddItem Or (enmTipoCampo = DB_BOOLEAN Or enmTipoCampo = DB_BYTE Or enmTipoCampo = DB_INTEGER Or enmTipoCampo = DB_LONG)
                    End If
                    If ((enmTipoExtra And extDecimal) <> 0) Then
                        bolAddItem = bolAddItem Or (enmTipoCampo = DB_CURRENCY Or enmTipoCampo = DB_QUANTITA Or enmTipoCampo = DB_DECIMAL Or enmTipoCampo = DB_DOUBLE Or enmTipoCampo = DB_SINGLE)
                    End If
                    If ((enmTipoExtra And extData) <> 0) Then
                        bolAddItem = bolAddItem Or (enmTipoCampo = DB_DATE Or enmTipoCampo = DB_DATETIME)
                    End If
                    If ((enmTipoExtra And extTime) <> 0) Then
                        bolAddItem = bolAddItem Or (enmTipoCampo = DB_TIME)
                    End If
                End If
                'aggiungo il campo al combo
                If bolAddItem Then
                    Call objCombo.addItem(Trim$(vetEntry(0)) & Space(20) & "#" & xAna.grinput(vetEntry(1)).NomeCampoDb)
                End If
            End If
NEXT_CaricaComboExtra:
        Loop While (strExtra <> "")
        Set xAna = Nothing
    End If
    
END_CaricaComboExtra:
    On Local Error GoTo 0
    CaricaComboExtra = (objCombo.ListCount > 0)
    Exit Function
    
ERR_CaricaComboExtra:
    Call MXNU.MsgBoxEX(1366, vbCritical, 1007, vetEntry(1))
    Resume NEXT_CaricaComboExtra
End Function

Public Sub SSExtraButtonClicked(MWAgt1 As MXKit.CAgenteAuto, AnagraficaExtra As MXKit.Anagrafica, SSExtra As FPSpreadADO.fpSpread, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim varNomeGruppo As Variant

    Call SSExtra.GetText(5, Row, varNomeGruppo)
    If AnagraficaExtra.EsisteGruppo(CStr(varNomeGruppo)) Then
        ' Rif. anomalia 5961
        If AnagraficaExtra.grinput(CStr(varNomeGruppo)).TipoValidazione <> "" Then
            Call AnagraficaExtra.TBselClick("", "", varNomeGruppo)
        End If
    Else
        If MXNU.ModuloRegole Then
            SSExtra.Col = Col
            SSExtra.Row = Row
            If SSExtra.CellType = SS_CELL_TYPE_BUTTON Then
                Dim objAgt As MXKit.CAgenteAuto
                
                Set objAgt = MXAA.CreaCAgenteAuto
                Set objAgt.formCorrente = MWAgt1.formCorrente
                Set objAgt.CurCtrl = SSExtra
                objAgt.Nome = ssCellGetValue(SSExtra, 2, Row)
                Call MXAA.EseguiAgt(objAgt)
                Set objAgt = Nothing
            End If
        End If
    End If

End Sub
Public Sub SSExtraEditMode(MWAgt1 As MXKit.CAgenteAuto, AnagraficaExtra As MXKit.Anagrafica, SSExtra As FPSpreadADO.fpSpread, ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    Dim vntValore As Variant, q%, vntNomeGruppo As Variant

    If MWAgt1.NoEditMode(SSExtra, Col&, Row&, Mode%, (ChangeMade)) Then Exit Sub
    If Mode = 0 And ChangeMade Then
        q = SSExtra.GetText(Col, Row, vntValore)
        q = SSExtra.GetText(5, Row, vntNomeGruppo)
        Call AnagraficaExtra.AssegnaCampo(CStr(vntNomeGruppo), vntValore)
    End If

End Sub

Private Function QueryDefinizione(strSezione As String, strNomeAnagrafica As String) As String
Dim strSQL As String
     Select Case UCase$(strSezione)
         Case "MAG"
             strSQL = "SELECT * FROM ExtraMag WHERE (CodArt = '')"
             strNomeAnagrafica = "ExtraMag"
         Case "DEPOSITI"
             strSQL = "SELECT * FROM ExtraDepositi WHERE (CodDeposito='')"
             strNomeAnagrafica = "ExtraDepositi"
         Case "DISTINTA"
             strSQL = "SELECT * FROM ExtraDistinta WHERE (Progressivo = 0)"
             strNomeAnagrafica = "ExtraDistinta"
         Case "AGENTI"
             strSQL = "SELECT * FROM ExtraAgenti WHERE (CodAgente = '')"
             strNomeAnagrafica = "ExtraAgenti"
         Case "BANCHE"
             strSQL = "SELECT * FROM ExtraBanche WHERE (CodBanca = '')"
             strNomeAnagrafica = "ExtraBanche"
         Case "CLIENTI"
             strSQL = "SELECT * FROM ExtraClienti WHERE (CodConto = '')"
             strNomeAnagrafica = "ExtraClienti"
         Case "FORNITORI"
             strSQL = "SELECT * FROM ExtraFornitori WHERE (CodConto = '')"
             strNomeAnagrafica = "ExtraFornitori"
         Case "GENERICI"
             strSQL = "SELECT * FROM ExtraGenerici WHERE (CodConto = '')"
             strNomeAnagrafica = "ExtraGenerici"
         Case "DIPENDENTI"
             strSQL = "SELECT * FROM ExtraDipendenti WHERE (CodDip='')"
             strNomeAnagrafica = "ExtraDipendenti"
         Case "CDLAVORO"
             strSQL = "SELECT * FROM ExtraCdLavoro WHERE (CodCdL='')"
             strNomeAnagrafica = "ExtraCdLavoro"
         Case "MACCHINE"
             strSQL = "SELECT * FROM ExtraMacchine WHERE (CodMac='')"
             strNomeAnagrafica = "ExtraMacchine"
         Case "CESPITI"
             strSQL = "SELECT * FROM ExtraCespiti WHERE (CodCespite = '')"
             strNomeAnagrafica = "ExtraCespiti"
         Case "COMMESSE"
             strSQL = "SELECT * FROM ExtraCommCli WHERE (RifProgressivo=0)"
             strNomeAnagrafica = "ExtraCommCli"
         Case "AVANZAMENTI"
             strSQL = "SELECT * FROM ExtraAvanzamenti WHERE (Movimento = 0)"
             strNomeAnagrafica = "ExtraAvanzamenti"
    End Select
    QueryDefinizione = strSQL
End Function

#If (ISM98SERVER <> 1) And (Estensione <> 1) Then
'definizione delle estensioni alle anagrafiche
Public Function DefEstensione(ByVal strNome As String, objLing As MWLinguetta, _
    objSch As MWSchedaBox, ByVal lngLingLeft As Long, _
    objAnag As Object, ctlExt As VBControlExtender, _
    Optional lngLeft As Long = 15, Optional lngTop As Long = 600, Optional ByVal strRiga As String = "") As Boolean
    
    Dim strFileIni As String ', strRiga As String
    Dim vntRiga As Variant
    Dim colObj As Collection
    Dim colAmb As Collection
    Dim NomeWrapper As String
    Dim ctlWrapper As VBControlExtender

    DoEvents
    DefEstensione = False
    On Local Error GoTo ERR_DefEstensione
    
#If ISNUCLEO <> 0 Then
    '*** Gestione dell'estensione su form anagrafiche da progetto nucleo  - 27 Maggio 2003
    If strRiga = "" Then
        strFileIni = CercaDirFile("DEBUGISV.INI", MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & ";" & MXNU.PercorsoPers)
        If strFileIni <> "" Then
            If UCase(Dir$(strFileIni, vbNormal)) = "DEBUGISV.INI" Then
                strRiga = MXNU.LeggiProfilo(strFileIni, "ESTENSIONI", strNome, "")
            End If
        End If
    End If
#End If
    
    If strRiga <> "" Then
        vntRiga = Split(strRiga, ";")
        On Local Error Resume Next
        objLing.Left = lngLingLeft
        objLing.Caption = MXNU.CaricaCaptionInLingua(vntRiga(0))
        objLing.Visible = True
        

        Call objSch.AggiungiLicenza(CStr(vntRiga(1)))
        On Local Error GoTo ERR_DefEstensione
        If UBound(vntRiga) >= 2 Then NomeWrapper = vntRiga(2) Else NomeWrapper = ""
        
        If NomeWrapper = "" Then
            Set ctlExt = objSch.ControlsEx.Add(vntRiga(1), OGGETTO_ESTENSIONE)
            ctlExt.Visible = True
            ctlExt.Left = lngLeft
            ctlExt.Top = lngTop
        Else
            On Local Error Resume Next
            Call objSch.AggiungiLicenza(NomeWrapper)
            On Local Error GoTo ERR_DefEstensione
            Set ctlWrapper = objSch.ControlsEx.Add(NomeWrapper, OGGETTO_WRAPPER_ESTENSIONE)
            Set ctlExt = ctlWrapper.object.CaricaEstensione(CStr(vntRiga(1)))
            ctlWrapper.Visible = True
            ctlWrapper.Left = lngLeft
            ctlWrapper.Top = lngTop
            ctlExt.Visible = True
        End If
        'ctlExt.Object.TabIndex = objSch.TabIndex
        Set objSch.Parent.FunzioniM98 = New CFunzioniMetodo98
        Set colObj = New Collection
        colObj.Add hndDBArchivi
        Set colAmb = Ambienti2Collection(MXNU.CheckOwner(ctlWrapper))
        DefEstensione = ctlExt.object.Inizializza(objSch.Parent, colAmb, colObj, objAnag)
        ' inizio rif.sch. A5660
        If DefEstensione Then
            If MXNU.MetodoXP Then
                Call ModificaLayoutControlli(ctlExt.object)
            End If
        Else
            If (NomeWrapper = "") Then
                Call objSch.ControlsEx.Remove(OGGETTO_ESTENSIONE)
            Else
                Call ctlWrapper.object.ScaricaEstensione
                Call objSch.ControlsEx.Remove(OGGETTO_WRAPPER_ESTENSIONE)
            End If
        End If
        ' fine rif.sch. A5660
    End If
    objLing.Visible = DefEstensione
    On Local Error GoTo 0
fine_DefEstensione:
    Set colObj = Nothing
    Set colAmb = Nothing
    Set ctlWrapper = Nothing
Exit Function
    
    
ERR_DefEstensione:
    Dim coderr&, dscerr$
    coderr = Err.Number
    dscerr = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbExclamation, strNome, Array("DefEstensione", coderr, dscerr))
    Resume fine_DefEstensione
    Resume
End Function
#End If

Public Function TerminaEstensione(ctlExt As VBControlExtender, ctlSch As MWSchedaBox) As Boolean
    
    Dim ctlWrapper As VBControlExtender
    
    If Not ctlExt Is Nothing Then
        ctlExt.Visible = False
        If Not ctlExt.object Is Nothing Then
            On Local Error Resume Next ' verifica presenza del wrapper
            Call ctlExt.object.Termina
            Set ctlSch.Parent.FunzioniM98 = Nothing
            
            Set ctlWrapper = ctlSch.ControlsEx(OGGETTO_WRAPPER_ESTENSIONE)
            If Err Then
                Set ctlWrapper = Nothing
                Err.Clear
            End If
            If ctlWrapper Is Nothing Then
                ctlSch.ControlsEx.Remove OGGETTO_ESTENSIONE
                Set ctlExt = Nothing
            Else
                ctlWrapper.Visible = False
                Call ctlWrapper.object.ScaricaEstensione
                ctlSch.ControlsEx.Remove OGGETTO_WRAPPER_ESTENSIONE
                Set ctlWrapper = Nothing
            End If
        End If
    End If
    
    TerminaEstensione = True
    
End Function




