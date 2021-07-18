Attribute VB_Name = "MMenu"
Option Explicit

'strumenti
Public Const ID_MNUSTRUMITEM_VINCOLI = 0
Public Const ID_MNUSTRUMITEM_REGFISC = 1
Public Const ID_MNUSTRUMITEM_CAMBIOUT = 4
Public Const ID_MNUSTRUMITEM_COLL = 6
Public Const ID_MNUSTRUMITEM_6_FIVE = 1
Public Const ID_MNUSTRUMITEM_LETTERE = 8
Public Const ID_MNUSTRUMITEM_11_SOLLECITI = 0
Public Const ID_MNUSTRUMITEM_11_RAGGSCAD = 1
Public Const ID_MNUSTRUMITEM_11_AVVCHDOC = 3
Public Const ID_MNUSTRUMITEM_PROGSTAMPA = 10

'costanti per esegui azione INI
Const ENTRY_AZIONE = "AZIONE"
Const AZIONE_AGENTE = "AGENTE"
Const AZIONE_ESTENSIONE = "ESTENSIONE"
Const AZIONE_EXE = "EXE"
Const AZIONE_FILTRO = "FILTRO"
Const AZIONE_FILTRO_TEMP = "FILTROTEMP"
Const AZIONE_SIT = "SITUAZIONE"
Const AZIONE_TAB = "TABELLA"
Const AZIONE_VIS = "VISIONE"
Const AZIONE_FILE = "FILE"
Const AZIONE_FILTRO_TAB = "FILTRO_TAB"
Const AZIONE_VIS_SEL = "VISIONE_SEL"
Const AZIONE_NAV = "NAVIGATORE"
Const AZIONE_ALLINONE = "ALLINONE"
Const AZIONE_QUALITY = "QUALITY"
Const AZIONE_SYNAPSE = "SYNAPSE"
Global mBolFaiControlloChiaveTSE As Boolean
Public Enum setStatoFinModuli
    SFM_INDEFINITO = 0
    SFM_NASCOSTA = 1
    SFM_VISIBILE = 2
    SFM_RIDIM = 3
End Enum

Sub AttivaMenuMetodo()

    If frmModuli.Stato = SFM_VISIBILE Then
        frmModuli.Stato = SFM_NASCOSTA
    Else
        frmModuli.Stato = SFM_VISIBILE
    End If

End Sub


Private Function ApriTSE() As Boolean
    Dim intRis As Integer
    Dim StrContr    As String
    On Local Error GoTo ApriTSE_ERR
    ApriTSE = True
    
    '******************************************************
    'MODIFICHE PER EVITARE IL CONTROLLO DELLA CHIAVE IN TSE
    If Not mBolFaiControlloChiaveTSE Then
        'Qui devo controllare sol se nella chiave è presente
        'il modulo 72 (licenza terminale server)
        'If Not MXNU.ControlloModulichiave(modtse) Then
        '    Call MXNU.MsgBoxEX(1624, vbOKOnly + vbCritical, 1007)
        '    ApriTSE = False
        'End If
        Exit Function
    End If
    
    'RIF.TECEUROLAB (09/09/2010)
    If (MXNU.ControlloModuliChiave(modProcessWatcher) = 0) Then
        ApriTSE = True
        Exit Function
    End If
    '******************************************************
    
    
    
    '>>> GESTIONE TSE
    '   collegamento al servizio TSE
    '[11/04/2011] Rimozione Chiave Hardware
'    intRis = MXNU.InitTSE()
'    If intRis Then
'        'richiede la stringa dei moduli disponibili al servizio TSE
'        intRis = MXNU.RichiediModuliTSE()
'    End If
'    If Not intRis Then
'        Call MXNU.MsgBoxEX(1624, vbOKOnly + vbCritical, 1007)
'    End If

ApriTSE_Fine:
    On Local Error GoTo 0
    ApriTSE = intRis
    Exit Function
    
ApriTSE_ERR:
    intRis = False
    Resume ApriTSE_Fine
End Function

'NOME           : CaricaModuliMetodo
'DESCRIZIONE    : carica la form dei moduli di metodo
Public Sub CaricaModuliMetodo(Optional bolCambioUtenteLingua As Boolean = False)
    
    '[11/04/2011] Rimozione Chiave Hardware
    'If ApriTSE() Then
        #If ISMETODO2005 = 1 Then
            If bolCambioUtenteLingua Then
                Dim bolPanelNascosto As Boolean
                bolPanelNascosto = metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden
                metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Close
                Call LoadIniMenu(frmModuli.TrwModuli)
                DoEvents
                Call frmModuli.CaricaPreferitiUtente
                If bolPanelNascosto Then
                    metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Closed = False
                    metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hide
                Else
                    metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Select
                End If
                '[11/04/2011] Rimozione Chiave Hardware
                'RIF.A#8750 - aggiorna i moduli e sblocca il file TSE
'                Call MXNU.AggiornaModuliTSE
            Else
                Load frmModuli2005    '<<< All'interno viene eseguito set frmModuli con l'istanza della frmMenu che contiene fisicamente l'albero dei moduli
            End If
        #Else
            #If ISMETODOXP = 1 Then
                If MXNU.MetodoXP Then
                    Set frmModuli = New frmModuliXP
                Else
                    Set frmModuli = New frmModuli98
                End If
            #Else
                Set frmModuli = New frmModuli98
            #End If
            Load frmModuli
            frmModuli.Stato = SFM_INDEFINITO
            frmModuli.Stato = SFM_NASCOSTA
            frmModuli.Show
            Call frmModuli.Ridimensiona
        #End If
        
    'Else
    '    Call ChiudiMetodo
    '    End
    'End If
        
End Sub



'esegue un comando standard associato ad un menù
Public Function EseguiAzioneINI(ByVal strNomeMenu As String, ByVal HelpContextID As Long, ByVal bolPers As Boolean) As Boolean

    Dim strAzione$
    Dim strFileMenu$
    Dim vntAzione As Variant
    Dim bolRes As Boolean
    Dim frmGenerica As Form
    Dim strPath As String
    Dim strFile As String
    Dim intq As Integer
    Dim CEsegui As Object
    Dim bolCaricaExt As Boolean
    
    bolRes = False
    On Local Error GoTo err_EseguiAzioneINI
'    If bolPers Then
'        strFileMenu = MXNU.File_ini_MenuPers
'    Else
        strFileMenu = MXNU.File_ini_Menu
'    End If
    DoEvents
    strAzione = MXNU.LeggiProfilo(strFileMenu, strNomeMenu, ENTRY_AZIONE, "")
    If strAzione <> "" Then
        'esecuzione dell'azione
        vntAzione = Split(strAzione, ";")
        Select Case UCase(vntAzione(0))
            Case AZIONE_AGENTE
                Call MXAA.EseguiAgt(Nothing, CStr(vntAzione(1)))
            
            Case AZIONE_ESTENSIONE
                
                ' inizio rif.svil. S1352
                If ControllaEstensioneAttiva(vntAzione) Then
                    bolCaricaExt = True
                    'Non possono essere caricate contemporaneamente la form documenti e documenti commerciali
                    If CStr(vntAzione(1)) = "TPDOC.ORC_EVOL" Then
                        bolCaricaExt = Not CaricataFormDocumenti
                        If Not bolCaricaExt Then
                            Call MXNU.MsgBoxEX(3067, vbCritical, 1007)
                        End If
                    End If
                    If bolCaricaExt Then
                        Set frmGenerica = New frmExtChild
                        frmGenerica.NomeEstensione = CStr(vntAzione(1))
                        frmGenerica.NomeWrapper = CStr(vntAzione(2))
                        Call FormLoader(frmGenerica, HelpContextID)
                        'Set frmGenerica = Nothing
                    End If
                End If
                ' fine rif.svil. S1352
            
            Case AZIONE_EXE
                Call EseguiShell(CStr(vntAzione(1)), Val(vntAzione(2)), Val(vntAzione(3)) <> 0)
            
            Case AZIONE_FILTRO
                If UBound(vntAzione) > 1 Then
                    Call PrintBlaster(UCase(vntAzione(1)), HelpContextID, vntAzione(2))
                Else
                    Call PrintBlaster(UCase(vntAzione(1)), HelpContextID)
                End If
            
            Case AZIONE_FILTRO_TEMP
                If UBound(vntAzione) > 1 Then
                    Call PrintBlaster(UCase(vntAzione(1)), HelpContextID, vntAzione(2), True)
                Else
                    Call PrintBlaster(UCase(vntAzione(1)), HelpContextID, , True)
                End If
            
            Case AZIONE_SIT
                Set frmGenerica = New frmSituazione
                Call frmGenerica.Situazione(CStr(vntAzione(1)), CStr(vntAzione(2)))
                Set frmGenerica = Nothing
            
            Case AZIONE_TAB
                Call CaricaTabella(CStr(vntAzione(1)), vntAzione(2), HelpContextID)
            
            Case AZIONE_VIS
                Set frmGenerica = New frmVisioni
                frmGenerica.HelpContextID = HelpContextID
                Call frmGenerica.Visione(CStr(vntAzione(1)), CStr(vntAzione(2)), CStr(vntAzione(3)))
                Set frmGenerica = Nothing
            
            Case AZIONE_FILE
                intq = InStrRev(CStr(vntAzione(1)), "\")
                If intq > 0 Then
                    strPath = Left(CStr(vntAzione(1)), intq)
                    strFile = Mid(CStr(vntAzione(1)), intq + 1)
                Else
                    strPath = ""
                    strFile = CStr(vntAzione(1))
                End If
                Call EseguiAppAssociata(strPath, strFile)
            
            Case AZIONE_FILTRO_TAB
                
                On Local Error Resume Next
                Set CEsegui = CreateObject(CStr(vntAzione(4)))
                If Err = 0 And Not CEsegui Is Nothing Then
                    Set CEsegui.XNU = MXNU
                    Set CEsegui.XDB = MXDB
                    Set frmGenerica = New frmFiltroTabella
                    frmGenerica.HelpContextID = HelpContextID
                    frmGenerica.IntCheckCol = 4
                    Call frmGenerica.Imposta(CEsegui, CStr(vntAzione(1)), CStr(vntAzione(2)), CStr(vntAzione(3)))
                    Set frmGenerica = Nothing
                End If
                Set CEsegui = Nothing

            '#################################################################################################################################################################################
            'Ripristinata gestione dell'azione Visione Con Selez. da Menu per poter gestire personalizzazioni per l'Amministrazione Metodo (es. Export Mov. Contabili verso Gamma Enterprise)
            'Attualmente abilitato solo per VisioniConSelezioni gestite da codice (tramite l'enum setTipoVisioneConSelez e la classe CVisioniConSelez dentro l'eseguibile).
            'Da valutare se abilitarlo verso l'esterno tramite l'oggetto CEsegui (vedere codice remmato qui sotto)
            '#################################################################################################################################################################################
            Case AZIONE_VIS_SEL
                On Local Error Resume Next
                Set frmGenerica = New frmVisioniConSelez
                frmGenerica.FormProp.FormID = HelpContextID
                Call frmGenerica.VisioneConSelez(CLng(vntAzione(1)), CStr(vntAzione(2)), "", "")
                Set frmGenerica = Nothing
            
'            Case AZIONE_VIS_SEL
'
'                On Local Error Resume Next
'                Set CEsegui = CreateObject(CStr(vntAzione(4)))
'                If Err = 0 And Not CEsegui Is Nothing Then
'                    Set CEsegui.XNU = MXNU
'                    Set CEsegui.XDB = MXDB
'                    Set frmGenerica = New frmVisioniConSelez
'                    frmGenerica.HelpContextID = HelpContextID
'                    Call frmGenerica.VisioneConSelez_Ext(CEsegui, CStr(vntAzione(1)), CStr(vntAzione(2)), CStr(vntAzione(3)))
'                    Set frmGenerica = Nothing
'                End If
'                Set CEsegui = Nothing

            Case AZIONE_NAV
                #If ISNUCLEO = 0 And TOOLS = 0 Then
                    Set frmGenerica = New frmNavigatore
                    frmGenerica.HelpContextID = HelpContextID
                    Call frmGenerica.Naviga(CStr(vntAzione(1)))
                #End If
            Case AZIONE_ALLINONE
                If ((MXNU.ControlloModuliChiave(modAllInOneRuntime) = 0) _
                    Or (MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0)) Then
                    'Dim frmDash As Form
                    Call MXALL.LoadDash(CStr(vntAzione(1)))
                    
                    'RIF.A#12100 - remmate righe successive. La chiamata MakeChild2 viene fatta dentro AIOT
                    'DoEvents
                    'Set frmDash = MXNU.FrmMetodo.FormAttiva
                    'MakeChild2 frmDash
                Else
                    Call MXNU.MsgBoxEX(9001, vbCritical, 1007, "[" & modAllInOneRuntime & "] All in One Touch (Runtime)")
                End If
            Case AZIONE_QUALITY
                Call MXQM.DoAction(CStr(vntAzione(1)), CStr(vntAzione(2)))
                DoEvents
                Dim frmQuality As Form
                Set frmQuality = MXNU.FrmMetodo.FormAttiva
                MakeChild2 frmQuality
            Case AZIONE_SYNAPSE
                If Not (NETFX Is Nothing) Then
                    Dim retVal As Variant
                    Dim dimpar As Integer
                    dimpar = UBound(vntAzione) - 1
                    Dim par() As String
                    ReDim par(dimpar)
                    Dim i As Integer
                    For i = 1 To UBound(vntAzione)
                        par(i - 1) = vntAzione(i)
                    Next
                    retVal = NETFX.RunSynapse(par)
                    
                End If
        End Select
        bolRes = True
    'Else
    '    bolres = True
    End If
    On Local Error GoTo 0
fine_EseguiAzioneINI:
    
    EseguiAzioneINI = bolRes
Exit Function

err_EseguiAzioneINI:
    Dim err_nr$
    Dim err_dsc$
    err_nr = Str(Err.Number)
    err_dsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("Esegui Azione (INI)", err_nr, err_dsc))
    Resume fine_EseguiAzioneINI
End Function

' rif.svil. S1352
' controlla se l'estensione è già stata lanciata
Private Function ControllaEstensioneAttiva(vntMenu As Variant) As Boolean

Dim intForm As Integer
Dim bolRes As Boolean
Dim strEst As String
        
    bolRes = True
    On Local Error GoTo END_ControllaEstensioneAttiva
    If UBound(vntMenu) > 2 Then
        If CStr(vntMenu(3)) = "1" Then
            strEst = UCase(CStr(vntMenu(1)))
            intForm = 0
            While (intForm < VB.Forms.Count) And (bolRes)
                If UCase(VB.Forms(intForm).NAME) = "FRMEXTCHILD" Then
                    If UCase(VB.Forms(intForm).NomeEstensione) = strEst Then
                        bolRes = False
                    End If
                End If
                If bolRes Then intForm = intForm + 1
            Wend
        End If
    End If
        
    
END_ControllaEstensioneAttiva:
    On Local Error Resume Next
    If Not (bolRes) Then
        If VB.Forms(intForm).WindowState = vbMinimized Then VB.Forms(intForm).WindowState = vbNormal
        VB.Forms(intForm).SetFocus
    End If
    On Local Error GoTo 0
    ControllaEstensioneAttiva = bolRes
    
End Function

Private Function CaricataFormDocumenti() As Boolean
    Dim frm As Form
    
    CaricataFormDocumenti = False
    On Local Error Resume Next
    For Each frm In VB.Forms
        If frm.NAME = "frmGestioneDoc" Then
            CaricataFormDocumenti = True
            Exit For
        End If
    Next

End Function




