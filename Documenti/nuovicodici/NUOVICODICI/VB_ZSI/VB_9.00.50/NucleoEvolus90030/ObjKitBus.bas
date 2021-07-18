Attribute VB_Name = "MObjKitBus"
Option Explicit

'dichiarazioni di metodo.bas condivise con mwserver
Global MXNU As MXNucleo.XNucleo
Global MXDB As MXKit.XODBC

Global hndDBArchivi As MXKit.CConnessione

Global MXCREP As MXKit.CAmbCRW
Global MXAA As MXKit.CAmbAgenti
Global MXCT As MXKit.CAmbTab
Global MXVI As MXKit.CAmbVisioni
Global MXVA As MXKit.CAmbValid
Global MXFT As MXKit.CAmbFiltri
Global MXWKF As MXKit.CAmbWorkFlow

Global MXSC As MXBusiness.CAmbScad
Global MXART As MXBusiness.CAmbVArt
Global MXSM As MXBusiness.CAmbStMag       'movimentazione storico magazzino
Global MXDBA As MXBusiness.CAmbDba        'gestione distinta base
Global MXGD As MXBusiness.CAmbGestDoc
Global MXPIAN As MXBusiness.CAmbPian
Global MXPN As MXBusiness.CAmbPN          'Prima Nota e Cespiti
Global MXPROD As MXBusiness.CAmbProd      'ambiente ordini di produzione
Global MXCICLI As MXBusiness.CAmbCicliLav 'ambiente cicli lavorazione
Global MXCC As MXBusiness.CAmbCommCli     'ambiente commesse clienti
Global MXRIS As MXBusiness.CAmbRisorse    'ambiente gestione risorse
Global MXSCH As MXBusiness.cAmbSched      'ambiente schedulazione
Global NETFX As Object 'ambiente dot net
Global GBolWorkflow As Boolean
'REMIND: modifiche per MXConsole
'Global MXALL As MXConsole.CAmbConsole
Global MXALL As Object

'REMIND: modifiche per Quality
Global MXQM As Object

'Modifiche per Wizard
Global MXWIZARD As Object

Private Enum setModuliRunTime
    MD32_KIT = 150
    MD32_BUSINESS_DBA = 160
    MD32_BUSINESS_PRIMANOTA = 161
    MD32_BUSINESS_SCADENZE = 162
    MD32_BUSINESS_STORICO = 163
    MD32_BUSINESS_DOCUMENTI = 164
    MD32_BUSINESS_PIANIFICAZIONE = 165
    MD32_BUSINESS_CTRLCODARTICOLO = 166
    MD32_BUSINESS_PRODUZIONE = 167
    MD32_BUSINESS_CICLILAVORAZIONE = 168
    MD32_BUSINESS_COMMESSECLIENTI = 169
    MD32_BUSINESS_GESTIONERISORSE = 170
    MD32_BUSINESS_SCHEDULAZIONE = 171
End Enum

'*** modifica ExtensionLoader ***
'ATTENZIONE: l'uso di una cache per la collezione ambienti crea una GROSSA falla di sicurezza. Agendo come segue, infatti
'non viene fatto il controllo sui moduli runtime isv:
'   1. Lancio un'estensione compilata da Metodo (Es. Estensione contatti su AnaCF)
'   2. Ambienti2Collection con bolSkipKey = true => mColAmb contiene TUTTI gli ambienti
'   3. Lancio un'estensione compilata da Rivenditore
'   4. Ambienti2Collection con bolSkipKey = false => dovrebbe fare il controllo dei moduli ISV
'   5. In realtà il controllo NON viene fatto perchè la funzione utilizza la cache fatta nel punto 2
'Private mColAmb As Collection

Public Function CreateObjKitBus(CTLXKit As Control, CTLXBus As Control) As Boolean

    CreateObjKitBus = True
    On Local Error GoTo CreateObjKitBus_Err
    
    If Not (CTLXKit Is Nothing) Then
        Set MXDB = CTLXKit.CreaXODBC()
        Set MXCREP = CTLXKit.CreaXCREP()
        Set MXVI = CTLXKit.CreaXVis()
        Set MXAA = CTLXKit.CreaXAgenti()
        Set MXCT = CTLXKit.CreaXTab()
        Set MXFT = CTLXKit.CreaXFT()
        Set MXVA = New MXKit.CAmbValid
        If GBolWorkflow Then
            Set MXWKF = CTLXKit.CreaWorkFlow
        End If
        
        
    End If
    If Not (CTLXBus Is Nothing) Then
        Set MXSC = CTLXBus.CreaXScad()
        Set MXART = CTLXBus.CreaXVArt()
        Set MXSM = CTLXBus.CreaXStMag()
        Set MXGD = CTLXBus.CreaXGestDoc()
        Set MXDBA = CTLXBus.CreaXDba()
        Set MXPIAN = CTLXBus.CreaXPianif()
        Set MXPN = CTLXBus.CreaXPrimaNota()
        Set MXPROD = CTLXBus.CreaXProduzione()
        Set MXCICLI = CTLXBus.CreaXCicliLavorazione()
        Set MXCC = CTLXBus.CreaXCommCli()
        Set MXRIS = CTLXBus.CreaXRisorse()
        Set MXSCH = CTLXBus.CreaXSchedulazione()
    End If



    'Rif. anomalia #7649
#If ISM98SERVER = 0 And ISTOOLS = 0 Then
        On Local Error Resume Next
        '[21/04/2011] Rimozione Chiave Hardware
        Set MXALL = CreateObject("MXConsole.CAmbConsole")
        Set MXQM = CreateObject("M98quality.cAmbQuality")
        Set MXWIZARD = CreateObject("MXWizard.cWizard")
        Set NETFX = CreateObject("MxHostNetFX.HostSynapseExecutor")
        
'        'REMIND: modifiche per MXConsole
'        'Set MXALL = New MXConsole.CAmbConsole
'        If ((MXNU.ControlloModulichiave(modAllInOneRuntime) = 0) _
'            Or MXNU.ControlloModulichiave(modMetodoXPEvolution) = 0) Then
'
'            Set MXALL = CreateObject("MXConsole.CAmbConsole")
'        End If
'
'        'REMIND: modifiche per Quality
'        If (MXNU.ControlloModulichiave(modQualityMenagement) = 0) Or (MXNU.ControlloModulichiave(modOfficeUser) = 0) Then
'            Set MXQM = CreateObject("M98quality.cAmbQuality")
'        End If
'
'        'Modifiche per Wizard
'        If (MXNU.ControlloModulichiave(modMetodoXPEvolution) = 0) Then
'            Set MXWIZARD = CreateObject("MXWizard.cWizard")
'        End If
        On Local Error GoTo CreateObjKitBus_Err
#End If

#If ISTOOLS <> 0 Then
    '[09/06/2011] Rimozione Chiave Hardware
    'If (MXNU.ControlloModulichiave(modMetodoXPEvolution) = 0) Then
        Set MXWIZARD = CreateObject("MXWizard.cWizard")
    'End If
#End If
    
CreateObjKitBus_Fine:
    On Local Error GoTo 0
    Exit Function
    
CreateObjKitBus_Err:
    Call MXNU.MsgBoxEX(9010, vbCritical, 1007, Array("CreateObjKitBus", Err.Number, Err.Description))
    CreateObjKitBus = False
    On Local Error GoTo 0
    Resume CreateObjKitBus_Fine
Resume
End Function

Public Function DropObjKitBus() As Boolean
Dim bolRes As Boolean

    bolRes = True
    
    'RIF.A#8908
    'Set mColAmb = Nothing
    
    'Rif. anomalia #7649
    #If ISM98SERVER = 0 Then
        If (Not MXWIZARD Is Nothing) Then
            Call MXWIZARD.Termina
            Set MXWIZARD = Nothing
        End If
    #End If
    
    'supporto scripting
    If Not MXAA Is Nothing Then MXAA.ResetAmbienti
    
    
    If Not MXSCH Is Nothing Then If MXSCH.Termina() Then Set MXSCH = Nothing Else bolRes = False
    If Not MXRIS Is Nothing Then If MXRIS.Termina() Then Set MXRIS = Nothing Else bolRes = False
    If Not MXCC Is Nothing Then If MXCC.Termina() Then Set MXCC = Nothing Else bolRes = False
    If Not MXCICLI Is Nothing Then If MXCICLI.Termina() Then Set MXCICLI = Nothing Else bolRes = False
    If Not MXPROD Is Nothing Then If MXPROD.Termina() Then Set MXPROD = Nothing Else bolRes = False
    If Not MXPIAN Is Nothing Then If MXPIAN.Termina() Then Set MXPIAN = Nothing Else bolRes = False
    If Not MXGD Is Nothing Then If MXGD.Termina() Then Set MXGD = Nothing Else bolRes = False
    If Not MXPN Is Nothing Then If MXPN.Termina() Then Set MXPN = Nothing Else bolRes = False
    If Not MXSM Is Nothing Then If MXSM.Termina() Then Set MXSM = Nothing Else bolRes = False
    If Not MXDBA Is Nothing Then If MXDBA.Termina() Then Set MXDBA = Nothing Else bolRes = False
    If Not MXART Is Nothing Then If MXART.Termina() Then Set MXART = Nothing Else bolRes = False
    If Not MXSC Is Nothing Then If MXSC.Termina() Then Set MXSC = Nothing Else bolRes = False
    If Not MXCT Is Nothing Then If MXCT.Termina() Then Set MXCT = Nothing Else bolRes = False
    If Not MXVA Is Nothing Then If MXVA.Termina() Then Set MXVA = Nothing Else bolRes = False
    If Not MXAA Is Nothing Then If MXAA.Termina() Then Set MXAA = Nothing Else bolRes = False
    If Not MXVI Is Nothing Then If MXVI.Termina() Then Set MXVI = Nothing Else bolRes = False
    If Not MXFT Is Nothing Then If MXFT.Termina() Then Set MXFT = Nothing Else bolRes = False
    If Not MXCREP Is Nothing Then If MXCREP.Termina() Then Set MXCREP = Nothing Else bolRes = False
    If Not NETFX Is Nothing Then If NETFX.Termina() Then Set NETFX = Nothing Else bolRes = False
    If Not MXWKF Is Nothing Then If MXWKF.Termina() Then Set MXWKF = Nothing Else bolRes = False
    
    'Rif. anomalia #7649
    #If ISM98SERVER <> 1 Then
        'REMIND: modifiche per MXConsole
        If (Not MXALL Is Nothing) Then
            Call MXALL.Terminate
            Set MXALL = Nothing
        End If
        'REMIND: modifiche per Quality
        If (Not MXQM Is Nothing) Then
            Call MXQM.Termina
            Set MXQM = Nothing
        End If
    #End If
    
    Set MXDB = Nothing
    Set MXNU = Nothing
    
    DropObjKitBus = bolRes
End Function

Public Function InitObjKitBus(hndDbArch As MXKit.CConnessione) As Boolean
    Dim bolWarning As Boolean
    Dim sLineErr As String

    InitObjKitBus = True
    bolWarning = False
    On Local Error GoTo InitObjKitBus_Err
    
    sLineErr = "INZIZIALIZZAZIONE INTERFACCIA CRYSTAL REPORTS"
    If Not (MXCREP Is Nothing) Then
        If Not MXCREP.Inizializza(MXNU) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Crystal Reports"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA FILTRI DI STAMPA"
    If Not (MXFT Is Nothing) Then
        If Not MXFT.Inizializza(MXNU, MXVI, MXDB, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Filtri di Stampa"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA VISIONI"
    If Not (MXVI Is Nothing) Then
        If Not MXVI.Inizializza(MXNU, MXDB, MXFT, MXCREP, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Visioni"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
   
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA AGENTI"
    If MXNU.ModuloRegole Then
        'Anomalia interna (inutile esposizione della proprietà ModuloRegole del nucleo in modifica/scrittura)
        ' La proprietà viene inizializzata in ChiavePresente() del nucleo e solo lì....
        'MXNU.ModuloRegole = MXAA.Inizializza(MXNU, MXDB, MXVI, MXCREP, hndDbArch) '<-- vecchia riga
        Call MXAA.Inizializza(MXNU, MXDB, MXVI, MXCREP, hndDbArch)
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''GESTIONE BPM''''''''''''''''''''''''''''''''''''''''
    If GBolWorkflow Then
        sLineErr = "INIZIALIZZAZIONE INTERFACCIA WORKFLOW"
        If Not (MXWKF Is Nothing) Then
            If Not MXWKF.Inizializza(MXNU, MXDB, hndDbArch) Then
                Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Workflow"))
                InitObjKitBus = False
                GoTo InitObjKitBus_Fine
            End If
        End If
        Dim loadwkfbusiness As New MXBusiness.LoadWorkFlow
        Dim loadwkfkit As New MXKit.LoadWorkFlow
        Call loadwkfbusiness.SetMXWKF(MXWKF)
        Call loadwkfkit.SetMXWKF(MXWKF)
        Set loadwkfbusiness = Nothing
        Set loadwkfkit = Nothing
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''GESTIONE BPM''''''''''''''''''''''''''''''''''''''''
     
#If ISMETODO2005 = 1 Then
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA NETFX"
      If Not (NETFX Is Nothing) Then
           Call NETFX.Inizializza(MXNU, mMetodoInterop)
    End If
#End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA VALIDAZIONI"
    If Not (MXVA Is Nothing) Then
        If Not MXVA.Inizializza(MXNU, MXDB, MXVI, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Validazioni"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA SCADENZE"
    If Not (MXSC Is Nothing) Then
        If Not MXSC.Inizializza(MXNU, MXDB, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Scadenze"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA TABELLE"
    If Not (MXCT Is Nothing) Then
        If Not MXCT.Inizializza(MXNU, MXDB, MXVI, MXAA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Tabelle"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA VALIDAZIONE ARTICOLI"
    If Not (MXART Is Nothing) Then
        If Not MXART.Inizializza(MXNU, MXDB, MXAA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Validazione Articoli"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA MOVIMENTAZIONE STORICO"
    If Not (MXSM Is Nothing) Then
        If Not MXSM.Inizializza(MXNU, MXDB, MXAA, MXART, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Movimentazione Magazzino"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA Prima Nota"
    If Not (MXPN Is Nothing) Then
        If Not MXPN.Inizializza(MXNU, MXDB, MXAA, MXCT, MXSC, MXVI, MXVA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Prima Nota"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA Documenti"
    If Not (MXGD Is Nothing) Then
        If Not MXGD.Inizializza(MXNU, MXDB, MXAA, MXART, MXSM, MXCT, MXSC, MXVI, MXPN, MXFT, MXCREP, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Gestione Documenti"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
#If ISTOOLS = 0 Then
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA DISTINTA BASE"
    If Not (MXDBA Is Nothing) Then
        If Not MXDBA.Inizializza(MXNU, MXDB, MXART, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Distinta Base"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
        
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA PIANIFICAZIONE"
    If Not (MXPIAN Is Nothing) Then
        If Not MXPIAN.Inizializza(MXNU, MXDB, MXART, MXDBA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Pianificazione"))
            InitObjKitBus = False
        End If
    End If
    
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA ORDINI DI PRODUZIONE"
    If Not (MXPROD Is Nothing) Then
        'RIF.A.ISV.#9 - aggiunto ambiente MXVA
        If Not MXPROD.Inizializza(MXNU, MXDB, MXAA, MXART, MXSM, MXCT, MXVI, MXDBA, MXPIAN, MXVA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Produzione"))
            InitObjKitBus = False
        End If
    End If
    
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA CICLI DI LAVORAZIONE"
    If Not (MXCICLI Is Nothing) Then
        If Not MXCICLI.Inizializza(MXNU, MXDB, MXART, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Cicli Lavorazione"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA COMMESSE CLIENTI"
    If Not (MXCC Is Nothing) Then
        If Not MXCC.Inizializza(MXNU, MXDB, MXAA, MXART, MXVI, MXDBA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Commesse Clienti"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAIZONE INTERFACCIA GESTIONE RISORSE"
    If Not (MXRIS Is Nothing) Then
        If Not MXRIS.Inizializza(MXNU, MXDB, MXAA, MXPROD, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Gestione Risorse"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAIZONE INTERFACCIA SCHEDULAZIONE"
    If Not (MXSCH Is Nothing) Then
        If Not MXSCH.Inizializza(MXNU, MXDB, MXAA, MXART, MXCT, MXVI, MXPROD, MXCICLI, MXRIS, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Schedulazione"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    bolWarning = True
       
    
    sLineErr = "INIZIALIZZAZIONE AMBIENTE ALLINONE"
    #If ISM98SERVER = 0 Then
        '[21/04/2011] Rimozione Chiave Hardware - Il controllo è stato spostato da CreateObjKitBus a InitObjKitBus
        If ((MXNU.ControlloModuliChiave(modAllInOneRuntime) = 0) _
            Or MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0) Then
            
            Dim colObjs As Collection
            Dim colAmbs As Collection
                    
            Set colAmbs = Ambienti2Collection(True)
            Set colObjs = New Collection
            colObjs.Add hndDBArchivi
            Call MXALL.Initialize(MXNU.PercorsoPgm & "\AllInOne", colAmbs, colObjs, False)
        Else
            Set MXALL = Nothing
        End If
    #End If

    sLineErr = "INIZIALIZZAZIONE AMBIENTE QUALITY"
    #If ISM98SERVER = 0 Then
        '[21/04/2011] Rimozione Chiave Hardware - Il controllo è stato spostato da CreateObjKitBus a InitObjKitBus
        If (MXNU.ControlloModuliChiave(modQualityMenagement) = 0) Or (MXNU.ControlloModuliChiave(modOfficeUser) = 0) Then
            If Not MXQM Is Nothing Then Call MXQM.Inizializza(MXNU)
        Else
            Set MXQM = Nothing
        End If
    #End If
#End If
 
    sLineErr = "INIZIALIZZAZIONE AMBIENTE WIZARD"
    #If ISM98SERVER = 0 Then
        '[21/04/2011] Rimozione Chiave Hardware - Il controllo è stato spostato da CreateObjKitBus a InitObjKitBus
        If (MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0) Then
            Call MXWIZARD.Inizializza(MXNU, MXDB, MXVI, MXVA, MXFT, MXCT, hndDBArchivi)
        Else
            Set MXWIZARD = Nothing
        End If
    #End If
   
    
InitObjKitBus_Fine:
    On Local Error GoTo 0
    Exit Function
    
InitObjKitBus_Err:
    Call MXNU.MsgBoxEX(9010, vbCritical, 1007, Array("InitObjKitBus", Err.Number, Err.Description & " [" & sLineErr & "]"))
    If Not bolWarning Then InitObjKitBus = False
    On Local Error GoTo 0
    Resume InitObjKitBus_Fine

End Function


Public Function Ambienti2Collection(Optional ByVal bolSkipKey As Boolean = False) As Collection
Dim colAmb As Collection

    'creo la collezione degli ambienti
    Set colAmb = New Collection
    With colAmb
        .Add MXNU, "MXNU"
        'NOTA: MXBROWSER non deve essere controllato da moduli RUNTIME ma dal solo modulo
        '      @METODO che viene controllato in fase di creazione dell'oggetto mMetodoBrowser
#If ISMETODO2005 = 1 Then
            .Add mMetodoBrowser, "MXBROWSER"
#End If
        
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_KIT) = 0) Then
            .Add MXDB, "MXDB"
            .Add MXCREP, "MXCREP"
            .Add MXCT, "MXCT"
            .Add MXVI, "MXVI"
            .Add MXVA, "MXVA"
            .Add MXFT, "MXFT"
            If MXNU.ControlloModuliChiave(modAgentiRunTime) = 0 Then .Add MXAA, "MXAA"
            .Add MXALL, "MXALL"
            .Add MXQM, "MXQM"
            .Add NETFX, "NETFX"
        End If
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_SCADENZE) = 0) Then .Add MXSC, "MXSC"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_CTRLCODARTICOLO) = 0) Then .Add MXART, "MXART"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_STORICO) = 0) Then .Add MXSM, "MXSM"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_DBA) = 0) Then .Add MXDBA, "MXDBA"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_DOCUMENTI) = 0) Then .Add MXGD, "MXGD"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_PIANIFICAZIONE) = 0) Then .Add MXPIAN, "MXPIAN"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_PRIMANOTA) = 0) Then .Add MXPN, "MXPN"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_PRODUZIONE) = 0) Then .Add MXPROD, "MXPROD"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_CICLILAVORAZIONE) = 0) Then .Add MXCICLI, "MXCICLI"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_COMMESSECLIENTI) = 0) Then .Add MXCC, "MXCC"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_GESTIONERISORSE) = 0) Then .Add MXRIS, "MXRIS"
    End With
    
    Set Ambienti2Collection = colAmb
    Set colAmb = Nothing

End Function

Public Function AddAmbienti2Script()
    
    With MXAA
        .AddAmbiente "MXRIS", MXRIS
        .AddAmbiente "MXCC", MXCC
        .AddAmbiente "MXCICLI", MXCICLI
        .AddAmbiente "MXPROD", MXPROD
        .AddAmbiente "MXPIAN", MXPIAN
        .AddAmbiente "MXGD", MXGD
        .AddAmbiente "MXPN", MXPN
        .AddAmbiente "MXSM", MXSM
        .AddAmbiente "MXDBA", MXDBA
        .AddAmbiente "MXART", MXART
        .AddAmbiente "MXSC", MXSC
        .AddAmbiente "MXCT", MXCT
        .AddAmbiente "MXVA", MXVA
        .AddAmbiente "MXVI", MXVI
        .AddAmbiente "MXFT", MXFT
        .AddAmbiente "MXCREP", MXCREP
        .AddAmbiente "NETFX", NETFX
        'aggiunta ambiente AIOT
        If (Not MXALL Is Nothing) Then
            .AddAmbiente "MXALL", MXALL
        End If
        
        '********* già presenti nella liberia ************
        '.AddAmbiente MXAA, "MXAA"
        '.AddAmbiente MXDB, "MXDB"
        '.AddAmbiente MXNU, "MXNU"
    End With
End Function

'funzione aggiunta per modulo acquisizione dati
Public Function Globals2Collection() As Collection
Dim colGlobs As Collection

    Set colGlobs = New Collection
    colGlobs.Add hndDBArchivi, "HNDDBARCHIVI"
    
    Set Globals2Collection = colGlobs
    Set colGlobs = Nothing
End Function

