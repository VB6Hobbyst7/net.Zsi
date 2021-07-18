Attribute VB_Name = "MGlobali"
Option Explicit
DefLng A-Z

Public MbolInChiusura As Boolean

Public Function CreaTempPerAnalisiDisp(strQueryVis As String, ByVal strWHERE As String, ByVal strOrderBy As String, ByVal vntDataElab As Variant) As Boolean
    Dim intPos As Integer
    Dim objAnalisiDisp As MXBusiness.CAnalisiProd
    
    MXNU.MostraMsgInfo 70035
    'creo oggetto per analisi
    Set objAnalisiDisp = MXPROD.CreaCAnalisiProd()
    CreaTempPerAnalisiDisp = Not (objAnalisiDisp Is Nothing)
    If CreaTempPerAnalisiDisp Then
        'effettuo analisi copertura
        CreaTempPerAnalisiDisp = objAnalisiDisp.AnalisiCopertura(strWHERE, vntDataElab)
        'modifico query visione
        intPos = InStrRev(strQueryVis, "where", , vbTextCompare)
        strQueryVis = Left$(strQueryVis, intPos - 2) & " WHERE IDSESSIONE=" & MXNU.IDSessione _
            & " ORDER BY " & strOrderBy
        Set objAnalisiDisp = Nothing
    End If
    MXNU.MostraMsgInfo ""
End Function


' validazioni personalizzate dei filtri
Public Sub ValidPersFiltri(ByVal strNomeValid As String, ByVal strNomeCmpValid As String, bolEseguiValStd As Boolean, vntNewValore As Variant)
    Select Case strNomeValid
        Case "VALID_ARTICOLO", "VALID_ARTVARIANTI"
            Dim xCodArt As MXBusiness.CVArt
            Set xCodArt = MXART.CreaCVArt()
            xCodArt.Codice = vntNewValore
            If xCodArt.Valida(CHIEDIVAR_TUTTE, False, , 0, False) Then
                vntNewValore = xCodArt.Codice
            End If
            'bolEseguiValStd = True
            Call xCodArt.Termina
            Set xCodArt = Nothing
        Case "VALID_ARTICOLOTIP"  ' rif.sch. A4562
            Dim xCodArtTip As MXBusiness.CVArt
            Set xCodArtTip = MXART.CreaCVArt()
            xCodArtTip.Codice = vntNewValore
            If xCodArtTip.Valida(CHIEDIVAR_TUTTE, False, , 0, False) Then
                vntNewValore = xCodArtTip.Codice
            End If
            'bolEseguiValStd = True
            Call xCodArtTip.Termina
            Set xCodArtTip = Nothing
        Case "VALID_ARTCOMPOSTO"
            Dim xCodDba As MXBusiness.CComposto
            Set xCodDba = MXDBA.CreaCComposto()
            If xCodDba.Valida(vntNewValore, False) Then
                vntNewValore = xCodDba.pCodice
            End If
            Set xCodDba = Nothing
        Case "VALID_CICLOPROD"
            Dim xCodClv As MXBusiness.CCicloLav
            Set xCodClv = MXCICLI.CreaCCiclo()
            If xCodClv.Valida(vntNewValore, False) Then
                vntNewValore = xCodClv.pCodice
            End If
            Set xCodClv = Nothing
    End Select
End Sub

Public Sub Totali_AddRecordIniziali(cTraccia As MXKit.cTraccia, _
                        HrsTot As MXKit.CRecordSet, _
                        bolRichiediParziali As Boolean, _
                        CmbTotali_ListIndex As Integer, _
                        IdRigaFiltro_Esecizio As Long, _
                        IdRigaFiltro_Data As Long, _
                        ssFiltroDati As Object, _
                        bolSituazione As Boolean, _
                        Optional vntCodConto As Variant, _
                        Optional IdRigaFiltro_CodConto As Variant, _
                        Optional IdRigaFiltro_Provisorio As Variant)
                        
    Dim objGruppo As MXKit.CGruppo
    Dim strQuery As String
    Dim strSelect As String
    Dim strWHERE As String
    Dim hSS As MXKit.CRecordSet
    Dim bolFinito As Boolean
    Dim vntValore As Variant
    Dim strNomeCampo As String
    Dim bolSaldoContabile As Boolean
    Dim vntEs As Variant
    Dim vntOldOpData As Variant
    
    'costruisco la query per i valori totali
    For Each objGruppo In cTraccia.pTotale(CmbTotali_ListIndex).colGruppi
        strNomeCampo = objGruppo.CColGruppo.strDataField
        If StrComp(cTraccia.pNomeTraccia, "VIS_MOVCON", vbTextCompare) = 0 Then
            If StrComp(strNomeCampo, "Mese", vbTextCompare) = 0 Then
                strNomeCampo = "0 as Mese"
                bolSaldoContabile = True
            End If
        End If
        strSelect = ConcatenaEspressione(strSelect, ",", strNomeCampo)
    Next objGruppo
    strQuery = "SELECT DISTINCT " & strSelect
    strQuery = strQuery & " FROM " & cTraccia.pLivelloCorrente.SQLDammiFROM
    vntOldOpData = ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Data))
    Call ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Data), OPMINUGUALE)
    If StrComp(cTraccia.pNomeTraccia, "VIS_MOVCON", vbTextCompare) = 0 Then
        Call ssFiltroDati.GetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntEs)
        Call ssFiltroDati.SetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntEs - 1)
    Else
        Call ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), OPMINUGUALE)
    End If
    cTraccia.pLivelloCorrente.strSQLWhr = cTraccia.CFiltroDati.SQLFiltro
    'Anomalia nr. 6310
    strWHERE = cTraccia.pLivelloCorrente.SQLDammiWHERE(False)
    Call ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), OPUGUALE)
    If StrComp(cTraccia.pNomeTraccia, "VIS_MOVMAG", vbTextCompare) = 0 Or StrComp(cTraccia.pNomeTraccia, "VIS_MOVMAG_BASE", vbTextCompare) = 0 Then
        Call ssCellLock(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio))
    End If
    If StrComp(cTraccia.pNomeTraccia, "VIS_MOVCON", vbTextCompare) = 0 Then
        Call ssFiltroDati.SetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntEs)
    End If
    Call ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Data), vntOldOpData)
    cTraccia.pLivelloCorrente.strSQLWhr = cTraccia.CFiltroDati.SQLFiltro
    strQuery = strQuery & " WHERE " & strWHERE
    'inserisco i dati nel recordset totali
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strQuery)
    bolFinito = MXDB.dbFineTab(hSS)
    If bolFinito And bolSaldoContabile And Not IsMissing(vntCodConto) Then
        'Situazione
        Call MXDB.dbInserisci(HrsTot)
        For Each objGruppo In cTraccia.pTotale(CmbTotali_ListIndex).colGruppi
            strNomeCampo = objGruppo.CColGruppo.strDataField
            If StrComp(strNomeCampo, "Esercizio", vbTextCompare) <> 0 Then
                Select Case UCase$(strNomeCampo)
                    Case "CONTO"
                        vntValore = vntCodConto
                    Case "MESE"
                        vntValore = "01-GEN"
                End Select
            Else
                Call ssFiltroDati.GetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntValore)
            End If
            Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, strNomeCampo, vntValore)
        Next objGruppo
        Call MXDB.dbRegistra(HrsTot)
    Else
        Do While Not bolFinito
            Call MXDB.dbInserisci(HrsTot)
            For Each objGruppo In cTraccia.pTotale(CmbTotali_ListIndex).colGruppi
                strNomeCampo = objGruppo.CColGruppo.strDataField
                If StrComp(strNomeCampo, "Esercizio", vbTextCompare) <> 0 Then
                    vntValore = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, strNomeCampo, "")
                    If StrComp(cTraccia.pNomeTraccia, "VIS_MOVCON", vbTextCompare) = 0 Then
                        If StrComp(strNomeCampo, "MESE", vbTextCompare) = 0 Then
                            vntValore = "01-GEN"
                        End If
                    End If
                Else
                    Call ssFiltroDati.GetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntValore)
                End If
                Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, strNomeCampo, vntValore)
            Next objGruppo
            Call MXDB.dbRegistra(HrsTot)
            bolFinito = Not MXDB.dbSuccessivo(hSS)
        Loop
    End If
    Call MXDB.dbChiudiSS(hSS)
    If bolSaldoContabile And IsMissing(vntCodConto) And Not IsMissing(IdRigaFiltro_CodConto) Then
        'Visione
        Dim strWHEIniz As String
        strWHEIniz = swapp(cTraccia.CFiltroDati.SQLFiltro(cTraccia.CFiltroDati.IdFiltro2Riga(Val(IdRigaFiltro_CodConto))), "VistaRigheContabilita.", "")
        If strWHEIniz <> "" Then strWHEIniz = strWHEIniz & " AND "
        strWHEIniz = strWHEIniz & cTraccia.CFiltroDati.SQLFiltro(cTraccia.CFiltroDati.IdFiltro2Riga(Val(IdRigaFiltro_Provisorio)))
        strQuery = "SELECT Conto FROM VistaSaldiInizialiPN WHERE Conto NOT IN (SELECT DISTINCT Conto FROM VistaRigheContabilita WHERE " & strWHERE & " ) AND " & strWHEIniz
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strQuery)
        bolFinito = MXDB.dbFineTab(hSS)
        Do While Not bolFinito
            Call MXDB.dbInserisci(HrsTot)
            
            Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, "Esercizio", vntEs)
            Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, "MESE", "01-GEN")
            
            vntValore = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Conto", "")
            Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, "Conto", vntValore)
           
            Call MXDB.dbRegistra(HrsTot)
            
            bolFinito = Not MXDB.dbSuccessivo(hSS)
        Loop
        Call MXDB.dbChiudiSS(hSS)
    End If
    bolRichiediParziali = False
    
End Sub


