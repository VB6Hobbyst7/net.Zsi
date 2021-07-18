VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CODEJOCK.COMMANDBARS.V15.3.1.OCX"
Begin VB.Form frmTabelle 
   Appearance      =   0  'Flat
   Caption         =   "Tabelle"
   ClientHeight    =   5400
   ClientLeft      =   2580
   ClientTop       =   2160
   ClientWidth     =   6660
   DrawStyle       =   6  'Inside Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "tabelle32.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   6660
   Begin FPSpreadADO.fpSpread Foglio 
      Height          =   3840
      Left            =   600
      TabIndex        =   0
      Top             =   900
      Width           =   5610
      _Version        =   524288
      _ExtentX        =   9895
      _ExtentY        =   6773
      _StockProps     =   64
      AutoCalc        =   0   'False
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   30
      NoBeep          =   -1  'True
      Protect         =   0   'False
      RestrictCols    =   -1  'True
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "tabelle32.frx":030A
      UnitType        =   2
      UserResize      =   1
      VisibleCols     =   5
      VisibleRows     =   14
      AppearanceStyle =   0
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   0
      Top             =   0
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTabelle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1

Public FormProp As New CFormProp
Public Maschera As Long
Public WithEvents cTab As MXKit.CTabelle
Attribute cTab.VB_VarHelpID = -1
Public MlngHlpTabella As Long
Public NOMETABELLA As String
Public DesTabella As String
Public ChiaveAgg As String
Public StrWheAgg As String

'Per Agenti Predefiniti
Dim AgenteTab$(1)
Dim MintAccessi As Integer
Public WithEvents MWAgt1 As MXKit.CAgenteAuto
Attribute MWAgt1.VB_VarHelpID = -1

Private MvetAgt() As String

Private Sub ctab_CancellazioneRiga(ByVal Row As Long, bolSuccesso As Boolean)
    Dim strSQL As String, hndSS As CRecordSet, intq As Integer, varCod As Variant
    Dim intCodSped As Integer
    bolSuccesso = True
    Select Case NOMETABELLA
        Case "TabSpediz"
            intq = Foglio.GetText(cTab.TTrovaColonna("Codice"), Row, varCod)
            strSQL = "SELECT Spedizioniere FROM TestaCostiSpedizione WHERE Spedizioniere = " & varCod
            Set hndSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
            If Not MXDB.dbFineTab(hndSS, TIPO_SNAPSHOT) Then
                If MXNU.MsgBoxEX(MXNU.CaricaStringaRes(1024, varCod), vbYesNo + vbDefaultButton2 + vbQuestion, MXNU.CaricaStringaRes(1007)) = vbNo Then
                    bolSuccesso = False
                End If
            End If
            intq = MXDB.dbChiudiSS(hndSS)
    End Select

End Sub

Private Sub CaricaFileSuCombo(strPath As String, strWildCard As String)
    'riempie il combo dei files contenuti nella directory indicata, controllando le wildcard.
    Dim objFSO As Scripting.FileSystemObject
    Dim objDir As Folder
    Dim objFile As file
    Dim objFiles As Files
    Dim i As Integer

    On Local Error GoTo CaricaFileSuCombo_ERR
    Set objFSO = New Scripting.FileSystemObject
    Call ssComboClear(Foglio, 3, -1)
    Call ssComboClear(Foglio, 4, -1)
    Call ssComboClear(Foglio, 5, -1)
    ReDim MvetAgt(0) As String
    If objFSO.FolderExists(strPath) Then
        Set objDir = objFSO.GetFolder(strPath)
        Set objFiles = objDir.Files
        For Each objFile In objFiles
            If UCase(Right(objFile.NAME, Len(strWildCard))) = strWildCard Then
                Call ssComboAddItem(Foglio, 3, -1, objFile.NAME)
                Call ssComboAddItem(Foglio, 4, -1, objFile.NAME)
                Call ssComboAddItem(Foglio, 5, -1, objFile.NAME)
                
                i = i + 1
                ReDim Preserve MvetAgt(0 To i) As String
                MvetAgt(i) = objFile.NAME
            End If
        Next
        Set objDir = Nothing
        Set objFile = Nothing
        Set objFiles = Nothing
    Else
        'Manca la directory
        Call MXNU.MsgBoxEX(MXNU.CaricaStringaRes(2231), vbOKOnly + vbExclamation, MXNU.CaricaStringaRes(2092))
    End If

FineSub:
    On Local Error GoTo 0
    Set objFSO = Nothing
    Set objDir = Nothing
    Set objFiles = Nothing
    Set objFile = Nothing
    Exit Sub
CaricaFileSuCombo_ERR:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Local Error Resume Next
    'Call MXDB.dbRollBack(hndDBArchivi)
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("CaricaFileSuCombo", errNumber, errDescription))
    Resume Next
End Sub

Private Sub CaricaAgenti()
    Dim lngr As Long
    Dim hSS As MXKit.CRecordSet
    
    For lngr = 1 To Foglio.DataRowCnt
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT * FROM TABCAT_ATI WHERE Codice=" & ssCellGetValue(Foglio, cTab.TTrovaColonna("Codice"), lngr))
        Call ssCellSetValue(Foglio, 3, lngr, TrovaElementoVet(MvetAgt, MXDB.dbGetCampo(hSS, NO_REPOSITION, "Agt_Tipo1", "")))
        Call ssCellSetValue(Foglio, 4, lngr, TrovaElementoVet(MvetAgt, MXDB.dbGetCampo(hSS, NO_REPOSITION, "Agt_Tipo2", "")))
        Call ssCellSetValue(Foglio, 5, lngr, TrovaElementoVet(MvetAgt, MXDB.dbGetCampo(hSS, NO_REPOSITION, "Agt_Tipo3", "")))
        Call MXDB.dbChiudiSS(hSS)
    Next lngr

End Sub

Private Sub ctab_DopoCaricamento()
    Dim varDato As Variant, lngi As Long, lngColonnaData As Long

    Select Case NOMETABELLA
        Case "TABCATAMM"
            Call LeggiTipiAmm
            Call CaricaFileSuCombo(MXNU.PercorsoPgm & "\Agenti\AmmTecnico\", ".CMP")
            Call CaricaAgenti
            
        Case "TabCausaliCes"
            For lngi = 1 To Foglio.DataRowCnt
                If ssCellGetValue(Foglio, 1, lngi) = 10 Or ssCellGetValue(Foglio, 1, lngi) = 14 Then
                    Call ssCellUnLock(Foglio, 3, lngi)
                Else
                    Call ssCellLock(Foglio, 3, lngi)
                End If
                If ssCellGetValue(Foglio, 1, lngi) = 10 Then
                    Call ssCellUnLock(Foglio, 4, lngi)
                Else
                    Call ssCellLock(Foglio, 4, lngi)
                End If
            Next lngi
        Case "ProgressiviStampa"
            lngColonnaData = cTab.TTrovaColonna("DataFin")
            For lngi = 1 To Foglio.DataRowCnt
                'varDato = Foglio.GetText(lngColonnaData, lngi, varDato)
                'If VarType(varDato) = vbEmpty Then
                varDato = ssCellGetValue(Foglio, lngColonnaData, lngi)
                If Not IsDate(varDato) Then
                    Call Foglio.SetText(lngColonnaData, lngi, MXNU.DataIniCont)
                End If
            Next lngi
            'Anomalia 11092
            Dim intTipoReg As Integer
            lngColonnaData = cTab.TTrovaColonna("DataFinRelIva")
            For lngi = 1 To Foglio.DataRowCnt
                varDato = ssCellGetValue(Foglio, lngColonnaData, lngi)
                If Not IsDate(varDato) Then
                    Call Foglio.SetText(lngColonnaData, lngi, MXNU.DataIniCont)
                End If
                intTipoReg = ssCellGetValue(Foglio, cTab.TTrovaColonna("TipoReg"), lngi)
                If Not (intTipoReg < 5 Or intTipoReg > 6) Then
                    ssCellLock Foglio, lngColonnaData, lngi
                    Call Foglio.SetText(lngColonnaData, lngi, "")
                End If
            Next lngi
        Case "TabImballi"
            Call cTab.CaricaDescrLingue(6)
        Case "TabPorto"
            Call cTab.CaricaDescrLingue(3)
        Case "TabTraspACura"
            Call cTab.CaricaDescrLingue(2)
        Case "TabCausTrasporto"
            Call cTab.CaricaDescrLingue(2)
        Case "TabEsenzioni"
            Call cTab.CaricaDescrLingue(4)
        Case "CategE98"
            Call cTab.CaricaDescrLingue(6)
        Case "MacroCategorieE98"
            Call cTab.CaricaDescrLingue(2)
        Case "TipoRegistroIva"
            Dim lngColonna As Long

            lngColonna = cTab.TTrovaColonna("Tappo")
            cTab.McolControlli(CStr(lngColonna)).StrWheAgg = "Esercizio=" & MXNU.AnnoAttivo
            lngColonna = cTab.TTrovaColonna("SelTappo")
            cTab.McolControlli(CStr(lngColonna)).StrWheAgg = "Esercizio=" & MXNU.AnnoAttivo
        'A #9720: bloccare la modifica del codice da VT02 a VT22
        Case "Regioni"
            For lngi = 1 To Foglio.DataRowCnt
                Select Case ssCellGetValue(Foglio, 1, lngi)
                    Case "VT02", "VT03", "VT04", "VT05", "VT06", "VT07", "VT08", "VT09", _
                        "VT10", "VT11", "VT12", "VT13", "VT14", "VT15", "VT16", "VT17", _
                        "VT18", "VT19", "VT20", "VT21", "VT22"
                        Call ssCellLock(Foglio, 1, lngi)
                End Select
            Next lngi
    End Select

End Sub

Private Sub LeggiTipiAmm()
    Dim hSS As MXKit.CRecordSet
    Dim i As Integer
    
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT * FROM TABTIPI_ATI ORDER BY Codice")
    While Not MXDB.dbFineTab(hSS)
        Call ssCellSetValue(Foglio, 3 + i, 0, MXNU.CaricaStringaRes(31958) & " " & MXDB.dbGetCampo(hSS, NO_REPOSITION, "Descrizione", ""))
        i = i + 1
        Call MXDB.dbSuccessivo(hSS)
    Wend
    Call MXDB.dbChiudiSS(hSS)
    
End Sub

Private Sub ctab_Registrazione(ByVal enmTipoReg As MXKit.setTipoRegistrazione, bolSuccesso As Boolean)
    Dim lngr As Long
    Dim strSQL As String
    Dim hSS As MXKit.CRecordSet
    Dim vntCodice As Variant
    Dim vntData As Variant
    
    bolSuccesso = True
    Select Case enmTipoReg
        Case vePrimaRegistrazione
            Select Case UCase(NOMETABELLA)
                Case "TABMASTRI", "TABCENTRICOSTO"
                    Call ImpostaDate
                Case "TIPOREGISTROIVA"
                    Dim intq As Integer, lngColonnaTipo As Long, varValore As Variant
                    Dim lngColonnaTappo As Long

                    bolSuccesso = True
                    lngColonnaTipo = cTab.TTrovaColonna("Tipo")
                    lngColonnaTappo = cTab.TTrovaColonna("Tappo")
                    For lngr = 1 To Foglio.DataRowCnt
                        intq = Foglio.GetText(lngColonnaTipo, lngr, varValore)
                        If InStr(UCase(varValore), "GIORNALE") <> 0 Then
                            Call Foglio.SetText(lngColonnaTappo, lngr, 0)
                        End If
                    Next lngr
                Case "TABFIRR"
                    Call MXDB.dbEseguiSQL(hndDBArchivi, "DELETE FROM TABFIRR")
                    For lngr = 1 To Foglio.DataRowCnt
                        Call ssCellSetValue(Foglio, cTab.TTrovaColonna("Progressivo"), lngr, lngr)
                        Call cTab.TrigaModificata(lngr)
                    Next lngr
                Case "TABLISTINI"
                    'Anomalia nr. 6216
                    Dim lngColonnaTipoL As Long
                    lngColonnaTipoL = cTab.TTrovaColonna("TP_TIPO")
                    For lngr = 1 To Foglio.DataRowCnt
                        If ssCellGetValue(Foglio, 0, lngr) = "M" And ssCellGetValue(Foglio, lngColonnaTipoL, lngr) = -1 Then
                            Call ssCellSetValue(Foglio, lngColonnaTipoL, lngr, 0)
                        End If
                    Next lngr
                Case "REGIONI"
                    'A #9720: bloccare l'eliminazione di un record da VT02 a VT22
                    For lngr = 1 To Foglio.DataRowCnt
                        If ssCellGetValue(Foglio, 0, lngr) = "A" Then
                            Select Case ssCellGetValue(Foglio, 1, lngr)
                                Case "VT02", "VT03", "VT04", "VT05", "VT06", "VT07", "VT08", "VT09", _
                                    "VT10", "VT11", "VT12", "VT13", "VT14", "VT15", "VT16", "VT17", _
                                    "VT18", "VT19", "VT20", "VT21", "VT22"
                                Call ssCellSetValue(Foglio, 0, lngr, "")
                                Call MXNU.MsgBoxEX(3183, vbCritical, 1007, Array(lngr, ssCellGetValue(Foglio, 1, lngr)))
                            End Select
                        End If
                    Next lngr
                Case "TABNAZIONI"
                    Dim objValid As New MXKit.ControlliCampo
                    Dim CodNazione As Variant
                    Dim bolChiesto As Boolean
                    Dim bolAggiorna As Boolean
                    
                    'Sviluppo 2964
                    objValid.Inizializza "VALID_NAZIONE"
                    objValid.ListaCampiRit = "FlgBlackList"
                    objValid.VisMessaggio = False
                    For lngr = 1 To Foglio.DataRowCnt
                        If ssCellGetValue(Foglio, 0, lngr) = "M" Then
                            CodNazione = ssCellGetValue(Foglio, cTab.TTrovaColonna("Codice"), lngr)
                            If objValid.Validazione(CodNazione) Then
                                If Cast(objValid.ValoriCampiRit("FlgBlackList"), vbInteger) <> Cast(ssCellGetValue(Foglio, cTab.TTrovaColonna("FlgBlackList"), lngr), vbInteger) Then
                                    If Not bolChiesto Then
                                        bolAggiorna = (MXNU.MsgBoxEX(3232, vbQuestion + vbYesNo, 1007) = vbYes)
                                        bolChiesto = True
                                    End If
                                    If bolAggiorna Then
                                        strSQL = "UPDATE ANAGRAFICACF SET FlgBlackList=abs(" & ssCellGetValue(Foglio, cTab.TTrovaColonna("FlgBlackList"), lngr) & ") WHERE CodNazione=" & hndDBArchivi.FormatoSQL(CodNazione, DB_DECIMAL)
                                        Call MXDB.dbEseguiSQL(hndDBArchivi, strSQL)
                                    End If
                                End If
                            End If
                        End If
                    Next
            End Select
        Case veDopoRegistrazione
            Select Case UCase(NOMETABELLA)
                Case "TABCATAMM"
                    For lngr = 1 To Foglio.DataRowCnt
                        Foglio.Row = lngr
                        Foglio.Col = 3
                        strSQL = "UPDATE TABCAT_ATI SET Agt_Tipo1='" & Foglio.text & "',"
                        Foglio.Col = 4
                        strSQL = strSQL & " Agt_Tipo2='" & Foglio.text & "',"
                        Foglio.Col = 5
                        strSQL = strSQL & " Agt_Tipo3='" & Foglio.text & "' " & _
                                 "WHERE Codice=" & ssCellGetValue(Foglio, cTab.TTrovaColonna("Codice"), lngr)
                        MXDB.dbEseguiSQL hndDBArchivi, strSQL
                    Next lngr
                    
                Case "TIPOREGISTROIVA"
                    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT Codice,DataIniCont FROM TabEsercizi")
                    intq = Not MXDB.dbFineTab(hSS)
                    While intq
                        vntCodice = MXDB.dbGetCampo(hSS, NO_REPOSITION, "Codice", 0)
                        vntData = CVDate(MXDB.dbGetCampo(hSS, NO_REPOSITION, "DataIniCont", "")) - 1
                        MXDB.dbEseguiSQL hndDBArchivi, "INSERT INTO ProgressiviStampa (Esercizio,NrRegistro,DataFin,DataFinRelIva,UtenteModifica,DataModifica) SELECT " & vntCodice & ",Codice,{d'" & Format$(vntData, "yyyy-mm-dd") & "'},{d'" & Format$(vntData, "yyyy-mm-dd") & "'},UtenteModifica,DataModifica FROM TipoRegistroIva WHERE Codice NOT IN (SELECT NrRegistro FROM ProgressiviStampa WHERE Esercizio= " & vntCodice & ")"
                        
                        intq = MXDB.dbSuccessivo(hSS)
                    Wend
                    intq = MXDB.dbChiudiSS(hSS)
                
                Case "PROGRESSIVISTAMPA"
                    For lngr = 1 To Foglio.DataRowCnt
                        vntCodice = ssCellGetValue(Foglio, cTab.TTrovaColonna("NRREGISTRO"), lngr)
                        vntData = ssCellGetValue(Foglio, cTab.TTrovaColonna("DATAFIN"), lngr)
                        If RegistroCespiti(vntCodice) Then
                            strSQL = "UPDATE PROGRESSIVISTAMPA SET DATAFIN=" & hndDBArchivi.FormatoSQL(vntData, DB_DATE) & _
                                     ",UTENTEMODIFICA=" & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT) & _
                                     ",DATAMODIFICA=" & hndDBArchivi.FormatoSQL(Now, DB_DATETIME) & _
                                     " WHERE NRREGISTRO=" & hndDBArchivi.FormatoSQL(vntCodice, DB_DECIMAL)
                            MXDB.dbEseguiSQL hndDBArchivi, strSQL
                        End If
                    Next lngr
            End Select
    End Select

End Sub

Private Function RegistroCespiti(vntCodice As Variant) As Boolean
    Dim hSS As MXKit.CRecordSet

    RegistroCespiti = False
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT Tipo FROM TipoRegistroIva WHERE Codice=" & hndDBArchivi.FormatoSQL(vntCodice, DB_DECIMAL))
    If Not MXDB.dbFineTab(hSS) Then
        RegistroCespiti = (MXDB.dbGetCampo(hSS, NO_REPOSITION, "Tipo", 0) = 8)
    End If
    Call MXDB.dbChiudiSS(hSS)
End Function


Private Sub ctab_SelezionePers(ByVal Col As Long, _
                               ByVal Row As Long, _
                               ByVal strCampoRif As String, _
                               ByVal strNomeValidazione As String, _
                               bolSuccesso As Boolean, _
                               bolEseguiValSt As Boolean, _
                               ByVal strListaCampi As String, _
                               colValoriRit As Collection)


End Sub

Private Sub ctab_ValidazionePers(ByVal strNomeCampo As String, _
                                 ByVal Row As Long, _
                                 ByVal enmTipoEvento As MXKit.SetTipoValidazione, _
                                 ByVal strNomeValidazione As String, _
                                 bolSuccesso As Boolean, _
                                 bolEseguiValSt As Boolean, _
                                 vntNewValore As Variant, _
                                 ByVal strListaCampi As String, _
                                 colValoriRit As Collection)

    bolSuccesso = False
    Select Case enmTipoEvento
        Case veValidazione
            Select Case UCase(NOMETABELLA)
                Case "TABSPEDIZ"
                    If UCase(strNomeCampo) = "PARTITAIVA" Then
                        'vntNewValore = Format(vntNewValore, "00000000000")   Anomalia nr. 12276
                        'Call ssCellSetValue(Foglio, cTab.TTrovaColonna("PARTITAIVA"), Row, vntNewValore)
                        Dim strValore As String
                        strValore = vntNewValore
                        Call MXVA.CtrlPartIVA(strValore, "", "", (ssCellGetValue(Foglio, cTab.TTrovaColonna("CODNAZIONE") + 3, Row)), True)
                        vntNewValore = strValore
                        Call ssCellSetValue(Foglio, cTab.TTrovaColonna("PARTITAIVA"), Row, vntNewValore)
                        bolEseguiValSt = False
                        bolSuccesso = True
                    End If
                Case "TABLINGUE"
                    If Val(vntNewValore) > 9 Then
                        bolSuccesso = False
                    Else
                        bolSuccesso = True
                    End If
                 Case "TABALIQUOTE"
                    If UCase(strNomeCampo) = "CODICE" Then
                        If vntNewValore > 99 Or vntNewValore = 0 Then   'Anomalia nr. 12287
                            bolSuccesso = False
                        Else
                            bolSuccesso = True
                        End If
                        bolEseguiValSt = False
                    End If
                Case "TABESENZIONI"
                    If UCase(strNomeCampo) = "CODICE" Then
                        If vntNewValore < 200 Or vntNewValore > 299 Then  'Rif. Anomalia Nr. 7469
                            bolSuccesso = False
                        Else
                            bolSuccesso = True
                        End If
                        bolEseguiValSt = False
                    End If
            End Select
        Case veCaricamento
            Select Case UCase(NOMETABELLA)
                Case "TABSPEDIZ"
                    If UCase(strNomeCampo) = "PARTITAIVA" Then
                        bolEseguiValSt = False
                        bolSuccesso = True
                    End If
            End Select
    End Select

End Sub

Private Sub Foglio_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    'rif.sch. A5820
    If UCase$(NOMETABELLA) = "TABUTENTI" Then
        If Col = 4 Then
            If ssCellGetValue(Foglio, Col, Row) > 32000 Then
                ' Msg: Il numero di terminale non può essere maggiore di 32000
                Call MXNU.MsgBoxEX(2786, vbInformation, 1007)
                Call ssCellSetValue(Foglio, Col, Row, 32000)
            End If
        End If
    ElseIf UCase(NOMETABELLA) = "TIPIEFFETTI" Then
        If Col <> cTab.TTrovaColonna("ModPagamentoPA") Then
            If ssComboListIndex(Foglio, cTab.TTrovaColonna("ModPagamentoPA"), Row) < 0 Then
                Call ssComboListIndex(Foglio, cTab.TTrovaColonna("ModPagamentoPA"), Row, 0)
            End If
        End If
        
    End If
    
End Sub

Private Sub Form_Activate()

    Call MXNU.Attiva_Toolbar(Me.hwnd, Maschera, Me.CommandBars)
    Call MXNU.ImpostaFormAttiva(Me)

End Sub

Private Sub Form_Load()
    Dim res%

    Me.HelpContextID = FormProp.FormID
    If Me.HelpContextID = 0 Then
        Me.HelpContextID = Me.MlngHlpTabella
    End If

    Call InitToolbarForm(Me)
    
    Set MWAgt1 = MXAA.CreaCAgenteAuto()
    
    If StrComp(NOMETABELLA, "TABLINGUE", vbTextCompare) = 0 Then
        Call InsertRecZero("TabLingue", 24125, "")
    End If
    If StrComp(NOMETABELLA, "TABNAZIONI", vbTextCompare) = 0 Then
        Call InsertRecZero("TabNazioni", 24971, "IT")
    End If
    
    Set cTab = MXCT.CreaCTabelle
    Call cTab.Inizializza(Foglio, MWAgt1)
    cTab.NOMETABELLA = Me.NOMETABELLA
    cTab.DesTabella = Me.DesTabella
    cTab.hWndForm = Me.hwnd
    cTab.ChiaveAgg = Me.ChiaveAgg
    cTab.StrWheAgg = Me.StrWheAgg
    DoEvents
    Call InizializzaTabelleParticolari

    Call cTab.TFormLoad(Me)
    
    'Anomalia nr. 5876
    If UCase(NOMETABELLA) = "TABLISTINI" Then
        'Anomalia nr. 10221, Sviluppo nr. 2944
        If Not ((MXNU.ControlloModuliChiave(modGDOGestioneListini) = 0) Or (MXNU.ControlloModuliChiave(modGDOImportExport) = 0)) Then
           Call ssColHide(Foglio, cTab.TTrovaColonna("TP_TIPO"))
        End If
    End If
    
    
    Select Case UCase(NOMETABELLA)
        'Sviluppo 1656
        Case "TABCATEGORIE", "TABCATEGORIESTAT", "CATEGORIECF", "TABSETTORI", "TABNAZIONI", "TABZONE"
            If Not MXNU.MOLAPChiavePresente() Then
                Call ssColHide(Foglio, cTab.TTrovaColonna("Budget"))
            End If
            'Sviluppo 2964
            If UCase(NOMETABELLA) = "TABNAZIONI" Then
                If MXNU.ControlloModuliChiave(modBlackList) <> 0 Then
                    Call ssColHide(Foglio, cTab.TTrovaColonna("FLGBLACKLIST"))
                End If
                Call cTab.CaricaDescrLingue(8)  'Anomalie nr. 7719, 10631, 11384
            End If
        'Fatturazione Elettronica P.A.
        Case "TIPIEFFETTI"
            If MXNU.ControlloModuliChiave(853) <> 0 Then
                Call ssColHide(Foglio, cTab.TTrovaColonna("MODPAGAMENTOPA"))
            End If
    End Select
    
    Maschera = BTN_TUTTI_MASK - BTN_MOD_MASK
    If cTab.FoglioFisso Then Maschera = Maschera - BTN_INS_MASK - BTN_ANN_MASK
    MintAccessi = FormImpostaAccessi(Me, Maschera)

    Call InizializzaCampiParticolari
    DoEvents
    res = MXAA.RegistraEventiFrm(Me, MWAgt1)
    Call MWAgt1.RegistraAgenteFrm(Me)
    
    If Not AccessoInserimento(MintAccessi) Then
        Foglio.MaxRows = Foglio.DataRowCnt
    End If
    
    'Anomalia nr. 5039
'Inzializzazione Form per Metodo Evolus
Call CambiaColoriControlli(Me)   'Da Mettere PRIMA di mResize
On Local Error Resume Next
Set mResize = New MxResizer.ResizerEngine
If (Not mResize Is Nothing) Then
        Call mResize.Initialize(Me, , , , , 0, True, metodo, MXNU)
End If
Call CentraFinestra(Me.hwnd)
Call CambiaCharSet(Me)
On Local Error GoTo 0
    Me.Show
    On Local Error Resume Next
    DoEvents
    MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled = True
    Me.CommandBars.FindControl(, idxBottoneTrova).Enabled = True
    On Local Error GoTo 0

End Sub

Private Sub Form_Paint()

    Call cTab.TFormPaint(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If AccessoModifica(MintAccessi) Then
        Call cTab.TFormQueryUnload(Me, Cancel, UnloadMode, AgenteTab())
    End If

'Per Metodo Evolus
If Not Cancel Then
        On Local Error Resume Next
        If (Not mResize Is Nothing) Then
                mResize.Terminate
                Set mResize = Nothing
        End If
        On Local Error GoTo 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call cTab.TFormUnload(Me, Cancel)
    Set FormProp = Nothing
    Set MWAgt1 = Nothing
    Set cTab = Nothing
    Set frmTabelle = Nothing

End Sub

Public Sub MetInserisci()

    Call cTab.TInserisci

End Sub


Public Function MetRegistra() As Boolean
    Dim Stato As Variant, intq As Integer

    If AccessoModifica(MintAccessi) Then
        intq = Foglio.GetText(0, (Foglio.ActiveRow), Stato)
        If Stato <> "A" Then Call cTab.TrigaModificata((Foglio.ActiveRow))
        intq = cTab.TsalvaTab(AgenteTab(), MWAgt1)
        MetRegistra = intq
    End If

End Function

Public Function AzioniMetodo(setAzione As enmFunzioniMetodo, Optional varparametro As Variant) As Variant

    Select Case setAzione
        Case MetFInserisci:         Call MetInserisci
        Case MetFRegistra:          AzioniMetodo = MetRegistra()
        Case MetFAnnulla:           Call MetAnnulla
        Case MetFPrecedente:        SendKeys "{UP}"
        Case MetFSuccessivo:        SendKeys "{DOWN}"
        Case MetFPrimo:             SendKeys "^{HOME}"
        Case MetFUltimo:            SendKeys "^{END}"
        Case MetFDettagli
        Case MetFStampa
        Case MetFVisUtenteModifica: Call MXCT.VisDatiUtenteModifica(FrmVisUtMod, "", "", "")
        Case MetFMostraCampiDBAnagr
            Dim ListaColl As New Collection
            ListaColl.Add cTab
            
            varparametro = False
            Set FrmNomiControlli.frmDef = Me
            Set FrmNomiControlli.ListaCol = ListaColl
            FrmNomiControlli.Show
            
            Set ListaColl = Nothing
        Case MetFDettVisione
        Case Else
    End Select
End Function
Public Sub MetAnnulla()

    Call cTab.TDelRec
    'Anomalia nr. 7995
    DoEvents
    Call ssCellActive(Foglio, Foglio.ActiveCol, Foglio.ActiveRow)

End Sub

Sub InizializzaTabelleParticolari()
    Dim lngr As Long
    
    Select Case UCase(NOMETABELLA)
        'Controlli particolari di Tabelle
        Case "TABSPEDIZ", "TABIMBALLI", "TABNAZIONI", _
             "TABPORTO", "TABTRASPACURA", "TABLINGUE", "TABCAUSTRASPORTO", _
             "TABMASTRI", "TABCENTRICOSTO", "TIPOREGISTROIVA", "CATEGE98", "TABFIRR", "MACROCATEGORIEE98", _
             "TABLISTINI", "TABCATAMM", "REGIONI"
             cTab.ControlloAggiuntivoTabella = True
        Case "PROGRESSIVISTAMPA"
             cTab.ControlloAggiuntivoTabella = True
             cTab.ChiaveAgg = "Esercizio," & DB_INTEGER & "," & MXNU.AnnoAttivo
        Case "TABCONTATORI"
             cTab.ChiaveAgg = "Esercizio," & DB_INTEGER & "," & MXNU.AnnoAttivo
        Case "ESITI"
             cTab.ValoreZeroAbilitato = True
        Case "TABCAUSALICES"
            cTab.ControlloAggiuntivoTabella = True
            For lngr = 1 To Foglio.DataRowCnt
                If lngr <> 10 And lngr <> 12 Then
                    Call ssDefInteger(Foglio, 4, lngr)
                End If
            Next lngr
    End Select

End Sub
Sub InizializzaCampiParticolari()
    Dim lngr As Long

    Select Case UCase(NOMETABELLA)
        'Controlli Particolari di campo
        Case "TABSPEDIZ"
            Call cTab.TSetControlloAggiuntivo("PARTITAIVA")
        Case "TABLINGUE"
            Call cTab.TSetControlloAggiuntivo("Codice")
        Case "TABCAUSALICES"
            For lngr = 1 To Foglio.DataRowCnt
                If lngr <> 10 Then
                    Call ssDefInteger(Foglio, 4, lngr)
                End If
            Next lngr
        Case "TABESENZIONI", "TABALIQUOTE"
            Call cTab.TSetControlloAggiuntivo("Codice")
    End Select

End Sub


Sub ImpostaDate()
    Dim lngr As Long, inti As Integer, varValoreCampo As Variant
    Dim lngColIniVal As Long, lngColFineVal As Long

    lngColIniVal = cTab.TTrovaColonna("DataIniValidita")
    lngColFineVal = cTab.TTrovaColonna("DataFineValidita")
    For lngr = 1 To Foglio.DataRowCnt
        inti = Foglio.GetText(0, lngr, varValoreCampo)
        If varValoreCampo <> "" Then
            inti = Foglio.GetText(lngColIniVal, lngr, varValoreCampo)
            If varValoreCampo = "" Then
                Call ssCellSetValue(Foglio, lngColIniVal, lngr, CVDate("1990/01/01"))
            End If
            inti = Foglio.GetText(lngColFineVal, lngr, varValoreCampo)
            If varValoreCampo = "" Then
                Call ssCellSetValue(Foglio, lngColFineVal, lngr, CVDate("2090/12/31"))
            End If
        End If
    Next lngr

End Sub

Private Sub InsertRecZero(strTabella As String, lngRisorsa As Long, strCodiceISO As String)
    Dim hSS As MXKit.CRecordSet
    Dim bolInserisci As Boolean
    Dim intq As Integer

    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT Codice FROM " & strTabella & " WHERE Codice=0")
    If MXDB.dbFineTab(hSS) Then
        bolInserisci = True
    End If
    intq = MXDB.dbChiudiSS(hSS)
    If bolInserisci Then
        If strCodiceISO <> "" Then
            Call MXDB.dbEseguiSQL(hndDBArchivi, "INSERT INTO " & strTabella & "(Codice,Descrizione,CodiceISO,CodStatoEstero,UtenteModifica,DataModifica) VALUES (0,'" & MXNU.CaricaStringaRes(lngRisorsa) & "','" & strCodiceISO & "',86,'" & MXNU.UtenteAttivo & "',{fn NOW()})")
        Else
            Call MXDB.dbEseguiSQL(hndDBArchivi, "INSERT INTO " & strTabella & "(Codice,Descrizione,UtenteModifica,DataModifica) VALUES (0,'" & MXNU.CaricaStringaRes(lngRisorsa) & "','" & MXNU.UtenteAttivo & "',{fn NOW()})")
        End If
    End If
        
End Sub

'Per Metodo Evolus
Private Sub mResize_AfterResize()
    Call AvvicinaLing(Me)
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call GestioneToolBut2005(Control.ID)
End Sub
