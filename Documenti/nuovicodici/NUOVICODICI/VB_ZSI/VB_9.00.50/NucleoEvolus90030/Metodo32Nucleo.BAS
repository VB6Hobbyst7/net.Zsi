Attribute VB_Name = "MMetodo"
Option Explicit
DefLng A-Z

Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetClipCursor Lib "user32" (lprc As RECT) As Long

Const MOUSEEVENTF_ABSOLUTE = &H8000   'spostamento assoluto
Const MOUSEEVENTF_LEFTDOWN = &H2    'pulsante sinistro premuto
Const MOUSEEVENTF_LEFTUP = &H4      'pulsante sinistro rilasciato

'=======================
'   tipi enumerativi
'=======================
Public Enum setBottoneAgente
    ageCollegamenti = 0
    ageMostraRiferimenti = 1
    ageNascondiRiferimenti = 2
    ageMostraCampiDB = 3
    ageNascondiCampiDB = 4
    ageDipendenze = 5
End Enum

Public Enum setTipoOpAn
    enmCompila = 0
    enmCarica = 1
End Enum

'=======================
'   costanti
'=======================

'VINCOLI GENERALI
'Global Const RS_MASTRO_CI = "MaCliIta"
'Global Const RS_MASTRO_CE = "MaCliEst"
'Global Const RS_MASTRO_FI = "MaForIta"
'Global Const RS_MASTRO_FE = "MaForEst"
Global Const RS_MASTRO_CLI = 1
Global Const RS_MASTRO_FOR = 2

Global Const GA_CreaDitta = 1
Global Const GA_CreaAnno = 2
Global Const GA_CopiaArchivi = 3
Global Const GA_CancellaAnno = 4
Global Const GA_TRASFSALDI = 5
Global Const GA_TRASFSCAD = 6
Global Const GA_TRASFPART = 7

Global Const SEL_DITTE = 0
Global Const SEL_ANNI = 1

'=============================================
'   dichiarazione costanti
'=============================================
Enum enmTestSalva
    tsnessuno = 0
    tsSalvato = 1
    tsNonSalvato = 2
    tsritorna = 3
End Enum
'=============================================
'=============================================
'   dichiarazione tipi di dati
'=============================================
'Gestione dei tasti della toolBox
'Type Metodo_Form_Attiva 'struttura che contiene informazioni sulla mdichild attiva
'     hwnd As Long   'handle
'     Tool_Mask As Long 'maschera dei tasti della toolbox
'End Type


'=============================================
'   dichiarazione variabili
'=============================================
'Global UltimoErr As Integer
'Global HlpAttivo As Integer
'Global SelAttiva  As Integer
'Global FormAttiva As Metodo_Form_Attiva

Global strinitexe As String
Global commitparziale As Integer
Global GTestMode As Boolean

'Globali per la stampa CRW
Global StpAVideo As Integer
'Global InStampa As Integer

Global hVinCfg As Integer      'handle del DYNASET dei S/Conti Generici Vincolati

Global frmModuli As Form

'*** DESIGNER ***
Global Designer As MXDesigner.cDesigner


'Flag per sapere se l'utente ha selezionato gli Extra Articoli o gli Extra Depositi
'per la Copia Archivi
Dim ExtraArtSel%
Dim ExtraDepSel%
Dim ExtraGiacDepSel%

'##### PER  METODO 2005 #################################################################
Dim MIdxBotAgentiAttuale As Long
Dim MIdxBotDesignerAttuale As Long
Dim MIdxBotZoomAttuale As Long
Dim MIdxBotTemaAttuale As Long

'per vedere se ci sono cambi ditta in corso
Global CmbDittaBusy As Boolean

#If ISMETODO2005 = 1 Then
  Global mMessagingEngine As Object
  Global mMetodoInterop As CMetodoInterop
  Global mMetodoBrowser As Object 'MxBrowser.CBrowserEngine
  Global GTemaAttivo As String
#End If
'########################################################################################

Global GBolNoMsgConfermaUscita As Boolean


#If ISMETODO2005 = 1 Then

    Public Sub InitToolbarForm(frmHost As Object)
        Dim tlbForm As XtremeCommandBars.CommandBar
        Dim Btn As CommandBarControl
        Dim btnPopup As CommandBarPopup
        
        'CommandBarsGlobalSettings.App = App
        'Risorse in lingua per i componenti Codejock
        CommandBarsGlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"
            
        If TypeOf frmHost Is Form Then
            frmHost.CommandBars.EnableCustomization False
            frmHost.CommandBars.ActiveMenuBar.Visible = False
            Set tlbForm = frmHost.CommandBars.Add("BarraForm", xtpBarTop)
            frmHost.CommandBars.Options.KeyboardCuesUse = xtpKeyboardCuesUseNone   'Anomalia 8691
            frmHost.CommandBars.KeyBindings.DeleteAll   'Necessario altrimenti in caso di caricamento di layout vecchi vengono reimpostate le combinazioni di tasti vecchie (es. F12 per aprire le Opzioni Generali)
        Else
            frmHost.Controls("CommandBars").EnableCustomization False
            frmHost.Controls("CommandBars").ActiveMenuBar.Visible = False
            Set tlbForm = frmHost.Controls("CommandBars").Add("BarraForm", xtpBarTop)
            frmHost.Controls("CommandBars").Options.KeyboardCuesUse = xtpKeyboardCuesUseNone   'Anomalia 8691
            frmHost.Controls("CommandBars").KeyBindings.DeleteAll   'Necessario altrimenti in caso di caricamento di layout vecchi vengono reimpostate le combinazioni di tasti vecchie (es. F12 per aprire le Opzioni Generali)
        End If
        tlbForm.BarID = ID_TLB_PRINC
        tlbForm.ContextMenuPresent = False
        
        tlbForm.Closeable = False
        tlbForm.Controls.DeleteAll
        
        With tlbForm.Controls
            Set Btn = .Add(xtpControlButton, idxBottoneInserisci, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "inserimento")
            Btn.ToolTipText = MXNU.CaricaStringaRes(2)
            Set Btn = .Add(xtpControlButton, idxBottoneDettaglio, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "dettagli")
            Btn.ToolTipText = MXNU.CaricaStringaRes(3)
            Set Btn = .Add(xtpControlButton, idxBottoneRegistra, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "registra")
            Btn.ToolTipText = MXNU.CaricaStringaRes(4)
            Set Btn = .Add(xtpControlButton, idxBottoneAnnulla, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "annulla")
            Btn.ToolTipText = MXNU.CaricaStringaRes(5)
            Set Btn = .Add(xtpControlButton, idxBottonePrimo, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "primo")
            Btn.ToolTipText = MXNU.CaricaStringaRes(6)
            Btn.BeginGroup = True
            Set Btn = .Add(xtpControlButton, idxBottonePrecedente, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "precedente")
            Btn.ToolTipText = MXNU.CaricaStringaRes(7)
            Set Btn = .Add(xtpControlButton, idxBottoneSuccessivo, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "succesivo")
            Btn.ToolTipText = MXNU.CaricaStringaRes(8)
            Set Btn = .Add(xtpControlButton, idxBottoneUltimo, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "ultimo")
            Btn.ToolTipText = MXNU.CaricaStringaRes(9)
            Set Btn = .Add(xtpControlButton, idxBottoneTrova, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "trova")
            Btn.ToolTipText = MXNU.CaricaStringaRes(10)
            Btn.BeginGroup = True
        End With
        tlbForm.ModifyStyle XTP_CBRS_GRIPPER, 0
        tlbForm.EnableDocking xtpFlagAlignTop
        If TypeOf frmHost Is Form Then
            frmHost.CommandBars.AddImageList metodo.ImglistBottoniXP
            If MXCtrl.TemaAttivo = "Office2007" Then
                frmHost.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
                frmHost.CommandBars.VisualTheme = xtpThemeResource
            ElseIf MXCtrl.TemaAttivo = "Office2010" Then
                frmHost.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
                frmHost.CommandBars.VisualTheme = xtpThemeRibbon    'xtpThemeResource
            Else
                frmHost.CommandBars.VisualTheme = MXNU.FrmMetodo.CommandBars.VisualTheme
            End If
            frmHost.CommandBars.Options.ShowExpandButtonAlways = False
            frmHost.CommandBars.DockToolBar tlbForm, 0, 0, xtpBarTop
        Else
            frmHost.Controls("CommandBars").AddImageList metodo.ImglistBottoniXP
            If MXCtrl.TemaAttivo = "Office2007" Then
                frmHost.Controls("CommandBars").GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
                frmHost.Controls("CommandBars").VisualTheme = xtpThemeResource
            ElseIf MXCtrl.TemaAttivo = "Office2010" Then
                frmHost.Controls("CommandBars").GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
                frmHost.Controls("CommandBars").VisualTheme = xtpThemeRibbon 'xtpThemeResource
            Else
                frmHost.Controls("CommandBars").VisualTheme = MXNU.FrmMetodo.CommandBars.VisualTheme
            End If
            frmHost.Controls("CommandBars").Options.ShowExpandButtonAlways = False
            frmHost.Controls("CommandBars").DockToolBar tlbForm, 0, 0, xtpBarTop
        End If
    End Sub

    Public Sub CambiaTema(NomeSkin As String)
        On Local Error Resume Next
        Dim bolPanelNascosto As Boolean
        metodo.MousePointer = vbHourglass
        bolPanelNascosto = metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden
        If Not bolPanelNascosto Then metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Close
        DoEvents
        
        If NomeSkin <> "" Then
            GTemaAttivo = NomeSkin  'Left(NomeSkin, InStrRev(NomeSkin, ".") - 1)
        Else
            GTemaAttivo = ""
        End If
        DoEvents
        Dim SysColor As OLE_COLOR
        
        CommandBarsGlobalSettings.ColorManager.EnableLunaBlueForRoyaleTheme = False
        DockingPaneGlobalSettings.ColorManager.EnableLunaBlueForRoyaleTheme = False
        SuiteControlsGlobalSettings.ColorManager.EnableLunaBlueForRoyaleTheme = False
        ShortcutBarGlobalSettings.ColorManager.EnableLunaBlueForRoyaleTheme = False
        
        Select Case LCase(GTemaAttivo)
            Case "office2010"
                SysColor = RGB(187, 206, 230)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
            Case "office2007"
                SysColor = RGB(191, 219, 255)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
            Case "winxp.luna"
                SysColor = RGB(230, 227, 210)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
            Case "winxp.royale"
                SysColor = RGB(228, 226, 230)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
            Case "vista"
                SysColor = RGB(232, 232, 232)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAero
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAero
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAero
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAero
            Case Else
                SysColor = 0
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
        End Select
        
        If InStr(1, NomeSkin, "Office2010", vbTextCompare) Then
            metodo.CommandBars.EnableOffice2007Frame True
            metodo.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
            metodo.CommandBars.VisualTheme = xtpThemeRibbon   'xtpThemeOffice2007
            DockingPaneGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
            metodo.DockingPaneManager.VisualTheme = ThemeResource
            ShortcutBarGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
            frmModuli2005.ShortcutBar1.VisualTheme = xtpShortcutThemeResource
            SuiteControlsGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
            CalendarGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
        ElseIf InStr(1, NomeSkin, "Office2007", vbTextCompare) Then
            metodo.CommandBars.EnableOffice2007Frame True
            metodo.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
            metodo.CommandBars.VisualTheme = xtpThemeResource
            DockingPaneGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
            metodo.DockingPaneManager.VisualTheme = ThemeResource
            ShortcutBarGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
            frmModuli2005.ShortcutBar1.VisualTheme = xtpShortcutThemeResource
            SuiteControlsGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
            CalendarGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
        ElseIf InStr(1, NomeSkin, "Luna", vbTextCompare) Then
            metodo.CommandBars.EnableOffice2007Frame False
            metodo.CommandBars.VisualTheme = xtpThemeOffice2003
            metodo.DockingPaneManager.VisualTheme = ThemeOffice2003
            frmModuli2005.ShortcutBar1.VisualTheme = xtpShortcutThemeOffice2003
        Else
            metodo.CommandBars.EnableOffice2007Frame False
            metodo.CommandBars.VisualTheme = xtpThemeVisualStudio2008
            metodo.DockingPaneManager.VisualTheme = ThemeExplorer
            frmModuli2005.ShortcutBar1.VisualTheme = xtpShortcutThemeNativeWinXP
        End If
        
        frmModuli.CommandBars.VisualTheme = metodo.CommandBars.VisualTheme
        frmModuli.ShortcutCaption1.VisualTheme = frmModuli2005.ShortcutBar1.VisualTheme
        frmModuli.ShortcutCaption2.VisualTheme = frmModuli2005.ShortcutBar1.VisualTheme
        
        'SetTreeViewBackColor frmModuli.TrwModuli, SysColor
        'SetTreeViewBackColor frmModuli.TrVSearch, SysColor
        'SetTreeViewBackColor frmModuli.TwPreferiti, SysColor
        
        frmMenu.MWSplitter1.BackColor = metodo.CommandBars.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
        If NomeSkin <> "" Then
            Call CambiaSchemaColori(True, SysColor, SysColor)
        Else
            Call CambiaSchemaColori(False, SysColor, SysColor)
        End If
        Call InizializzaSpread(True, True, SysGradientColor1)
        'Inizializzo lo spread anche per MXKit e per MXBusiness
        Call MXVI.InitSpreadEvolus
        Call MXBusiness.InitSpreadEvolus
        DoEvents
        Call metodo.CambiaColoriFormAttive(True)
        DoEvents
        
        If Not bolPanelNascosto Then metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Select
        
        DoEvents
        metodo.MousePointer = vbDefault
        On Local Error GoTo 0
        MXCtrl.TemaAttivo = GTemaAttivo
        
Uscita:
        Exit Sub
        
'CtrlSkinMenus:
'    Dim oDocument As MSXML2.DOMDocument
'    Dim UserNode As MSXML2.IXMLDOMNode
'    Dim objNode As MSXML2.IXMLDOMNode
'
'    On Local Error Resume Next
'    'In alcuni sistemi operativi (es. Windows XP) applicando la skin anche al menu non funziona più l'attivazione della barra dei menu della MDI tramite tastiera (tasto ALT oppure F10)
'    'Leggo se attivare la skin dal file di configurazione
'    If Dir$(MXNU.PercorsoPreferenze & "\SkinFrameworkExclude.xml", vbNormal) <> "" Then
'        Set oDocument = New MSXML2.DOMDocument
'        If oDocument.Load(MXNU.PercorsoPreferenze & "\SkinFrameworkExclude.xml") Then
'            Set UserNode = oDocument.selectSingleNode("//user[@name='" & LCase(MXNU.NomeComputer) & "']")
'            If (UserNode Is Nothing) Then
'                Set UserNode = oDocument.selectSingleNode("//user[@name='*']")
'            End If
'            If Not (UserNode Is Nothing) Then
'                Set objNode = UserNode.selectSingleNode("skinmenus")
'                If Not (objNode Is Nothing) Then
'                    If objNode.text <> "0" Then
'                        metodo.SkinFramework.ApplyOptions = metodo.SkinFramework.ApplyOptions Or xtpSkinApplyMenus
'                    End If
'                End If
'            End If
'        End If
'        Set objNode = Nothing
'        Set UserNode = Nothing
'        Set oDocument = Nothing
'    End If
'Return
    End Sub
   
    Public Sub GestioneToolBut2005(ByVal btnId As Long)
    
        Dim frmAttiva As Form
        Dim bolCancellaAzione As Boolean
        Dim intAzioneEseguita As Integer
        Dim bolApriNomiCtrl As Boolean
        Dim bolAllInOne As Boolean
        'Dim DesignButton As MSComctlLib.Button ' DESIGNER
        Dim DesignButton As XtremeCommandBars.CommandBarControl
        Dim frmDesign As Form
        
        Dim intForm As Integer
        Dim bolFrmZoomAperta As Boolean
        Dim objExt As Object
        Dim Button As Object  'CommandBarButton
        Dim bolChkEnabled As Boolean, bolValido As Boolean
        
        bolChkEnabled = True
        Select Case btnId
            Case idxBottoneAttivaDesigner To idxBottoneGrigliaDesigner
                Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneDesigner).CommandBar.FindControl(, btnId)
            '**** Bottone Zoom tolto su 9.00.00 *******************************************************
            'Case idxBottoneZoom100 To idxBottoneZoom200
            '    Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneZoom).CommandBar.FindControl(, btnId)
            Case idxBottoneDesigner
                If MIdxBotDesignerAttuale = 0 Then MIdxBotDesignerAttuale = idxBottoneAttivaDesigner
                Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneDesigner).CommandBar.FindControl(, MIdxBotDesignerAttuale)
                btnId = MIdxBotDesignerAttuale
            '**** Bottone Zoom tolto su 9.00.00 *******************************************************
            'Case idxBottoneZoom
            '    If MIdxBotZoomAttuale = 0 Then MIdxBotZoomAttuale = idxBottoneZoom100
            '    Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneZoom).CommandBar.FindControl(, MIdxBotZoomAttuale)
            '    btnId = MIdxBotZoomAttuale
            Case ID_TLBITEM_CHANGETHEME
                If MIdxBotTemaAttuale = 0 Then MIdxBotTemaAttuale = ID_TLBITEM_THEME_OFFICE2010   'ID_TLBITEM_THEME_OFFICE2007
                bolChkEnabled = False
            Case ID_TLBITEM_THEME_OFFICE2007 To ID_TLBITEM_THEME_VISTA, ID_TLBITEM_THEME_SYSTEM, ID_TLBITEM_THEME_OFFICE2010
                Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, btnId)
            Case Else
                Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, btnId)
        End Select
        If bolChkEnabled Then
            If Not (Button Is Nothing) Then
                bolValido = Button.Enabled
            Else
                bolValido = False
            End If
        Else
            bolValido = True
        End If
        'If Button.Enabled Then
        If bolValido Then
            On Local Error Resume Next
            Set frmAttiva = metodo.FormAttiva
            metodo!helpTimer.Interval = 500
            '***If TypeName(Button) = "IButtonMenu" Then
                Select Case btnId
                ' Gestione Zoom su foglio scheda sviluppo n.ro 1128 + Designer
                '***If metodo.Barra.Buttons(idxBottoneDefAgenti).Value = tbrUnpressed And LCase$(Left$(Button.key, 4)) <> "zoom" And LCase$(Left$(Button.key, 6)) <> "design" Then
                '*** Select Case Button.Index
                        Case idxBottoneDefAgenti    'Def. Agenti
                            ' Rif. anomalia n.ro 4824
                            If MXNU.ModuloRegole And Not (MXNU.CtrlAccessi) Then
                                Call FormImpostaAgenti(frmAttiva)
                            End If
                        Case idxBottoneNomiCtrlCmp  'Nome Controlli - campi
                            If (metodo.FormsCount > 2) Then
                                bolApriNomiCtrl = True
                                Call frmAttiva.AzioniMetodo(MetFMostraCampiDBAnagr, bolApriNomiCtrl)
                                If bolApriNomiCtrl Then
                                    Set FrmNomiControlli.frmDef = frmAttiva
                                    FrmNomiControlli.Show
                                End If
                            End If
                        Case idxBottoneSituazAnagr   'Situazione Anagrafica
                            'mostro le dipendenze dell'anagrafica
                            Dim ListaCol As New Collection
                            Call frmAttiva.AzioniMetodo(MetFVisDipendenze, ListaCol)
                            If ListaCol.Count > 0 Then
                                Call MXVA.VisualizzaDipendenze(ListaCol)
                            End If
                            Set ListaCol = Nothing
                        Case idxBottoneAttivaFileLog   'Attiva\Disattiva Log
                            'If Button.Caption = "Disattiva File Log" Then
                            If Button.Checked Then
                                Call MXDB.DisattivaLog
                                Button.ToolTipText = "Attiva File Log"
                                Call frmLog.MostraFileLog(MXNU.GetTempDir & "MWSqlLog" & MXNU.NTerminale & ".sql")
                                Button.Checked = False
                            Else
                                Call MXDB.AttivaLog
                                Button.ToolTipText = "Disattiva File Log"
                                Button.Checked = True
                            End If
                        Case idxBottoneRicProfili   'ricarica profili
                            metodo.MousePointer = vbHourglass
                            MXNU.MostraMsgInfo 70058
                            
                            '>>> profili anagrafiche
                            If Not MXVA Is Nothing Then
                                If MXVA.ChiudiDyTRAnagraf() Then
                                    Call MXVA.ApriDyTRAnagraf
                                End If
                                If MXVA.ChiudiDyTRValidazione() Then
                                    Call MXVA.ApriDyTRValidazione
                                End If
                            End If
                            '>>> profili tabelle
                            If Not MXCT Is Nothing Then
                                If MXCT.ChiudiDyTRTabelle() Then
                                    Call MXCT.ApriDyTRTabelle
                                End If
                            End If
                            '>>> profili visioni/situazioni
                            If Not MXVI Is Nothing Then
                                If MXVI.ChiudiDyTRSituazioni() Then
                                    Call MXVI.ApriDyTRSituazioni
                                End If
                                If MXVI.ChiudiDyTRVisioni() Then
                                    Call MXVI.ApriDyTRVisioni
                                End If
                            End If
                            MXNU.MostraMsgInfo ""
                            metodo.MousePointer = vbDefault
                            Button.DefaultItem = True
                            
                            '>>> forzo ricaricamente regole synapse
                            Dim synexe As Object
                            If (Not NETFX Is Nothing) Then
                                Set synexe = NETFX.GetHostSynapse
                                synexe.Dispose
                                synexe = Nothing
                            End If
                '***ElseIf LCase$(Left$(Button.key, 6)) = "design" Then ' DESIGNER
                        Case idxBottoneAttivaDesigner To idxBottoneGrigliaDesigner
                '***End If
            '***Else
                Err.Clear
                    Case ID_TLBITEM_FRMORIGINALSIZE
                        If frmAttiva.Name <> "EmptyForm" And Not EsisteElementoCollection(McolFormsInNavBar, frmAttiva.Name) Then
                            Call frmAttiva.Frm_SetOriginalSize
                            If Err.Number <> 0 Then
                                Err.Clear
                                Call frmAttiva.mResize.Frm_SetOriginalSize
                            End If
                        End If
                    Case idxBottoneModuli
                        If metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden Then
                            'il menu moduli non e' aperto ne lockato quindi lo apro e lo locko
                            metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Close
                            metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Select
                            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, btnId).Checked = True
                            If frmModuli.ShortcutBar1.Selected.ID = ID_BAR_PROGMODULES Then
                                Call DaiFocusAlberoModuli
                            End If
                        Else
                            'il menu moduli e' gia' aperto e lockato, lo chiudo
                            metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hide
                            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, btnId).Checked = False
                        End If
                    
                    Case idxBottoneAllInOne
                    Case idxBottoneInserisci
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_insert", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFInserisci))
                            'call frmAttiva.AzioniMetodo(MetFInserisci)
                            Call EseguiAgenteAzioneForm(frmAttiva, "after_insert", bolCancellaAzione, intAzioneEseguita)
                        End If
                    Case idxBottoneDettaglio
                        Call DoLostFocus(frmAttiva.ActiveControl)
                        Call frmAttiva.AzioniMetodo(MetFDettagli)
                        
                    Case idxBottoneRegistra
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_save", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFRegistra))
                            Call EseguiAgenteAzioneForm(frmAttiva, "after_save", bolCancellaAzione, intAzioneEseguita)
                        End If
                                    
                    Case idxBottoneAnnulla
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_delete", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFAnnulla))
                            'Call frmAttiva.AzioniMetodo(MetFAnnulla)
                            Call EseguiAgenteAzioneForm(frmAttiva, "after_delete", bolCancellaAzione, intAzioneEseguita)
                        End If
                    
                    Case idxBottonePrimo
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_first", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            Call frmAttiva.AzioniMetodo(MetFPrimo)
                        End If
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_first", bolCancellaAzione)
                    
                    Case idxBottonePrecedente
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_previous", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            Call frmAttiva.AzioniMetodo(MetFPrecedente)
                        End If
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_previous", bolCancellaAzione)
                    
                    Case idxBottoneSuccessivo
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_next", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            Call frmAttiva.AzioniMetodo(MetFSuccessivo)
                        End If
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_next", bolCancellaAzione)
                    Case idxBottoneUltimo
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_last", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            Call frmAttiva.AzioniMetodo(MetFUltimo)
                        End If
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_last", bolCancellaAzione)
                    Case idxBottoneStampa
                        Call frmAttiva.AzioniMetodo(MetFStampa)
                    Case idxBottoneTrova
                        If Not MXCT.TTrova(FrmTrovaGen) Then
                            Call frmAttiva.AzioniMetodo(MetFTrova)
                        End If
                    ' S#3040 - rimossa la Gestione Accessi da Evolus
'                    Case idxBottoneDefAccessi
'                        Call FormDefinisciAccessi(frmAttiva)
                    Case idxBottoneHelp
                        '#If ISMETODOXP = 1 Then
                        '    If MXNU.MetodoXP = True Then
                                    #If TOOLS = 1 Then
                                        If MXNU.DammiFormAttiva.Name = "frmModuli" Then
                                            Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
                                        Else
                                            Call MXNU.ApriHelp(False)
                                        End If
                                    #Else
                                        'posso caricare l'help direttamente su il browser della form tooltip
                                        If MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "METODOW", "HelpCompilato", 0) = 1 Then
                                            Call MXNU.ApriHelp(False)
                                        End If
                                    #End If
                            'Else
                            '    #If TOOLS = 1 Then
                            '        If MXNU.DammiFormAttiva.Name = "frmModuli" Then
                            '            Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
                            '        Else
                            '            Call MXNU.ApriHelp(False)
                            '        End If
                            '    #Else
                            '        Call MXNU.ApriHelp(False)
                            '    #End If
                            'End If
                        '#Else
                        '    #If TOOLS = 1 Then
                        '        If MXNU.DammiFormAttiva.Name = "frmModuli" Then
                        '            Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
                        '        Else
                        '            Call MXNU.ApriHelp(False)
                        '        End If
                        '    #Else
                        '        Call MXNU.ApriHelp(False)
                        '    #End If
                        '#End If
                    Case idxBottoneUtenteModifica
                        Call frmAttiva.AzioniMetodo(MetFVisUtenteModifica)
                    Case idxBottoneSchedula
                        Call frmAttiva.AzioniMetodo(MetFSchedulaOperazione)
                    'Case idxBottoneSpostaBarra
                    '    metodo!Barra.Align = (metodo!Barra.Align) Mod 2 + 1
                    '    metodo!Barra.Buttons(Button.Index).Image = metodo!Barra.Align + 10
                    '    metodo!tlbDesigner.Align = (metodo!tlbDesigner.Align) Mod 2 + 1 ' Designer
'**** Bottone Zoom tolto su 8.04.00 **************************************************************************************************
'                    Case idxBottoneZoom
'                        '#If ISMETODOXP = 1 Then
'                        '  If MXNU.MetodoXP Then
'                            Dim intZoomType As Integer
'                            intZoomType = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\Mw.ini", MXNU.UtenteAttivo, "ZoomType", 1)
'                            ' Gestione dell'apertura di una singola form di zoom
'                            bolFrmZoomAperta = False
'                            intForm = 0
'                            While (intForm < VB.Forms.Count) And (Not (bolFrmZoomAperta))
'                                If UCase(VB.Forms(intForm).NAME) = "FRMZOOM" Then
'                                    bolFrmZoomAperta = True
'                                End If
'                                intForm = intForm + 1
'                            Wend
'
'                            If Not bolFrmZoomAperta Then
'                                ' Rif. anomalia ISV n.ro 6 / Anomalia Metodo XP n.ro 6084 / Quesito ISV 294
'                                If LCase(frmAttiva.NAME) = "frmextchild" Then
'                                    If frmAttiva.ActiveControl.object.Controls.Count > 1 Then
'                                        Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl, intZoomType, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl.ActiveCol, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
'                                    Else
'                                        Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl, intZoomType, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl.ActiveCol, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
'                                    End If
'                                ElseIf TypeName(frmAttiva.ActiveControl) = "MWSchedaBox" Then
'                                    ' Sono in una scheda
'                                    If frmAttiva.ActiveControl.ControlsEx.Count > 0 Then
'                                        ' La scheda contiene delle estensioni
'                                        ' Setto lo UserControl dentro il wrapper in objExt (ExtWrapper deve avere una public property get come tutti gli UserControl estensioni)
'                                        Set objExt = frmAttiva.ActiveControl.ControlsEx(0).object.Controls
'                                        'Rif. anomalia #8088
'                                        If TypeName(objExt(0)) = "CTLXBus" Then
'                                            Set objActiveControl = objExt(1).object.ExtActiveControl
'                                        Else
'                                            Set objActiveControl = objExt(0).object.ExtActiveControl
'                                        End If
'                                        Set objExt = Nothing
'                                        Call frmZoom.DoZoom(frmAttiva, objActiveControl, intZoomType, objActiveControl.ActiveCol, objActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
'                                        Set objActiveControl = Nothing
'                                    End If
'                                Else
'                                    Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl, intZoomType, frmAttiva.ActiveControl.ActiveCol, frmAttiva.ActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
'                                End If
'                            Else
'                                Call MXNU.MsgBoxEX(2738, vbExclamation, 23481)
'                            End If
'
'                        '  End If
'                        '#End If
'*****************************************************************************************************************************************
                    Case idxBottoneDesigner
                    Case ID_TLBITEM_INFO
                      Call ShellExecute(metodo.hwnd, "Open", "www.metodo.it/MetodoPortal/DesktopPortalMain.aspx?action=NEWS", "", App.Path, 1)
                    Case ID_TLBITEM_PREFERENCES
                    Case ID_TLBITEM_THEME_OFFICE2007 To ID_TLBITEM_THEME_VISTA, ID_TLBITEM_THEME_SYSTEM, ID_TLBITEM_THEME_OFFICE2010
                        Dim NomeFileTema As String
                        Select Case btnId
                            Case ID_TLBITEM_THEME_OFFICE2010: NomeFileTema = "Office2010"
                            Case ID_TLBITEM_THEME_OFFICE2007: NomeFileTema = "Office2007"
                            Case ID_TLBITEM_THEME_WINXPLUNA: NomeFileTema = "WinXP.Luna"
                            Case ID_TLBITEM_THEME_WINXPROYALE: NomeFileTema = "WinXP.Royale"
                            Case ID_TLBITEM_THEME_VISTA: NomeFileTema = "Vista"
                            Case ID_TLBITEM_THEME_SYSTEM: NomeFileTema = ""
                        End Select
                        Call CambiaTema(NomeFileTema)
                        

                        If MIdxBotTemaAttuale = 0 Then
                            Dim i%
                            For i = ID_TLBITEM_THEME_OFFICE2007 To ID_TLBITEM_THEME_VISTA
                                metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, i).DefaultItem = False
                            Next i
                            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2010).DefaultItem = False
                        Else
                            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, MIdxBotTemaAttuale).DefaultItem = False
                        End If
                        metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, btnId).DefaultItem = True
                        MIdxBotTemaAttuale = btnId
                End Select
            '***End If
            Set frmAttiva = Nothing
        End If
    
    End Sub
    
    'Nel caso si docki la navigation bar, il focus non è posizionato sull'albero dei moduli per un problema degli oggetti Codejock;
    'aggiro il problema simulando un click del mouse.
    Public Sub DaiFocusAlberoModuli()
        Dim pt As POINTAPI, OldPt As POINTAPI
        Dim oldClipCur As RECT
        Dim r As RECT
        Dim bolCalcolaPt As Boolean
        On Local Error Resume Next
        
        Call GetCursorPos(OldPt)   'Memorizzo l'attuale posizione del mouse
        Call GetClipCursor(oldClipCur)
        DoEvents
        Sleep 30   'Attenzione: Sleep necessario altrimenti và in errore di protezione generale!!
        
        'Determino le coordinate di un punto all'interno dell'albero dei moduli...
        pt.x = (frmMenu.TrwModuli.Left + frmMenu.TrwModuli.Width / 2) \ Screen.TwipsPerPixelX
        'pt.y = (frmMenu.TrwModuli.Top + frmMenu.TrwModuli.Height / 2) \ Screen.TwipsPerPixelY
        pt.y = 150
        
        'Confino il cursore all'interno dell'albero dei moduli
        Call GetWindowRect(frmMenu.TrwModuli.hwnd, r)
        Call ClipCursor(r)
        
        '...e ci posiziono il mouse
        SetCursorPos pt.x, pt.y
        DoEvents
        'Simulo il click del mouse
        mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP, pt.x, pt.y, 0, 0
        DoEvents
        'Riposiziono il mouse alle coordinate originali
        Call ClipCursor(oldClipCur)
        
        SetCursorPos OldPt.x, OldPt.y
        On Local Error GoTo 0
    End Sub
    
    
#End If


Sub AbilitaDisabilitaMetodo(ByVal abilita As Boolean)
   Dim q%, ad%
   'Salva la Maschera dei Bottoni Attivi quando premo il bottone delle Regole
   If Not abilita Then
        MXNU.MascheraAttiva = 0
        Call MXNU.DammiMascheraAttiva
   End If
   'Abilita/Diabilita tutte le Form Attive
   For q = 0 To Forms.Count - 1
        Forms(q).Enabled = abilita
   Next q
   metodo.Enabled = True
   'Abilita/Disabilita tutti i controlli della form Principale
    For q = 0 To metodo.Controls.Count - 1
        If TypeOf metodo.Controls(q) Is Menu Then
            If metodo.Controls(q).Caption <> "-" Then metodo.Controls(q).Enabled = abilita
        ElseIf TypeOf metodo.Controls(q) Is ToolBar Then
            If metodo.Controls(q).Index <> 19 Then
                metodo.Controls(q).Enabled = abilita
            End If
        End If
    Next q
    'Ripristina i Bottoni Attivi quando premo per la seconda volta il bottone delle Regole
    If abilita Then
        Call MXNU.SetMascheraAttiva
    End If
End Sub


Sub ChiudiFormAttive()
    Dim q As Integer
    Dim strNome As String
    Dim intLimitemax As Integer
    Dim bolUnload As Boolean
        
    #If ISMETODO2005 = 1 Then
        If MbolInChiusura Then
            intLimitemax = 1
        Else
            intLimitemax = McolFormsInNavBar.Count + 1
        End If
    #Else
        'Rif.sch. 8198 - (form di log con error 91)
        If MbolInChiusura Then
            intLimitemax = 1
        Else
            intLimitemax = 2
        End If
    #End If
    Do
        If Forms.Count <= intLimitemax Then
            Exit Do
        Else
            For q = 0 To Forms.Count - 1
                strNome = Forms(q).Name
                bolUnload = False
                If (StrComp(strNome, "frmModuli98", vbTextCompare) <> 0) _
                    And (StrComp(strNome, "frmModuliXP", vbTextCompare) <> 0) _
                    And (StrComp(strNome, "metodo", vbTextCompare) <> 0) _
                    And (StrComp(strNome, "frmToolTip", vbTextCompare) <> 0) Then
                    bolUnload = True
                    #If ISMETODO2005 Then
                        If Not MbolInChiusura Then
                            bolUnload = Not EsisteElementoCollection(McolFormsInNavBar, strNome)
                        End If
                    #End If
                    If bolUnload Then
                        Dim hwnd As Long
                        hwnd = Forms(q).hwnd
                        Unload Forms(q)
                        If IsWindow(hwnd) Then CmbDittaBusy = True
                        While IsWindow(hwnd)
                            DoEvents
                            Sleep 100
                        Wend
                        CmbDittaBusy = False
                        Exit For
                    End If
                End If
            Next q
        End If
    Loop

'    q = 1
'    Do While Forms.Count > intLimitemax
'        strNome = Forms(Forms.Count - q).Name
'        'Rif. Anomalie98 Nr. 3536: la frmModuli ha cambiato nome; è frmModuli98 oppure frmModuliXP
'        If (StrComp(strNome, "frmModuli98", vbTextCompare) <> 0) _
'            And (StrComp(strNome, "frmModuliXP", vbTextCompare) <> 0) _
'            And (StrComp(strNome, "metodo", vbTextCompare) <> 0) _
'            And (StrComp(strNome, "frmToolTip", vbTextCompare) <> 0) Then
'            Unload Forms(Forms.Count - q)
'            q = 1
'        Else
'            If StrComp(strNome, "frmModuli98", vbTextCompare) = 0 _
'               Or StrComp(strNome, "frmModuliXP", vbTextCompare) = 0 _
'               Or StrComp(strNome, "frmToolTip", vbTextCompare) = 0 Then
'                q = 2
'            ElseIf StrComp(strNome, "metodo", vbTextCompare) = 0 Then
'                q = 3
'            End If
'        End If
'    Loop
End Sub

'Chiude tutte i database attivi,tutte le form ed esce dal programma.
Sub ChiudiMetodo()
    Dim q As Integer

    If VB.Forms.Count > 1 Then
        If Not (frmSelDitta Is Nothing) Then
            Unload frmSelDitta
            Set frmSelDitta = Nothing
        End If
    End If
    
    #If USAM98SERVER Then
    If Not GobjM98Server Is Nothing Then
        'Call GobjM98Server.Termina
        Set GobjM98Server = Nothing
    End If
    #End If
    
    'On Local Error Resume Next
    'chiusura recordset calendari
    If MXNU.CalendariLocale Then
        MXCICLI.TerminaRecSetCalendari
    End If
    'On Local Error GoTo 0
    
    'RIF.A#7602 - Si assicura che anche tutte le form di AIOT siano chiuse
    If (Not MXALL Is Nothing) Then
        Call MXALL.CloseAllWindows
    End If
    
    If Not (hndDBArchivi Is Nothing) Then q = MXDB.dbChiudiDB(hndDBArchivi)
    q = MXDB.dbDisattiva()

    Call DropObjKitBus
    
#If ISMETODO2005 = 1 Then
    On Local Error Resume Next
    'chiusura mailing system
    If (Not mMessagingEngine Is Nothing) Then
        mMessagingEngine.Dispose
        Set mMessagingEngine = Nothing
    End If
    
    'distruzione dell'hosting Metodo
    Set mMetodoInterop = Nothing
#End If
    
    'While RegolaBloccata()
    '    DoEvents
    'Wend
    'ChiudiRegola
    'EndAmbiente

End Sub

Sub AggiornaStatusBar()
    Dim strStato As String

    'aggiorno la caption dell'applicazione
    #If TOOLS = 1 Then
        metodo.Caption = App.Title & " - [" & MXNU.PercorsoPreferenze & "]"
    #Else
        metodo.Caption = App.Title & IIf(MXNU.ControlloModuliChiave(modChiaveDemo) = 0, " - Versione dimostrativa non destinata alla vendita", vbNullString)
    #End If
    'aggiorno la barra di stato
    Select Case MXNU.StatoEsercizioCont
'        Case 0, 2
'            strStato = " [T]"
'        Case 1
'            strStato = ""
'        Case 3
'            strStato = " [C]"
        Case 0
            strStato = ""
        Case 1
            strStato = " [C]"
    End Select
    metodo.BarraStato.Panels("dittaanno").Text = MXNU.Dsc_Breve_Ditta & " - " & MXNU.AnnoAttivo & " " & strStato
    metodo.BarraStato.Panels("utente").Text = MXNU.UtenteAttivo & " (" & MXNU.NTerminale & ")"
    metodo.BarraStato.Panels("Lingua").Text = MXNU.LinguaAttiva
    #If ISMETODOXP = 1 Then
        If MXNU.MetodoXP Then
            metodo.BarraStato.Panels("Designer").Text = MXNU.VersioneAttiva
        End If
    #End If
    #If ISMETODO2005 = 1 Then
        metodo.BarraStato.Panels("dittaanno").Width = 0  'AutoSize
        metodo.BarraStato.Panels("dittaanno").Width = metodo.BarraStato.Panels("dittaanno").Width + 20
        metodo.BarraStato.Panels("utente").Width = 0  'AutoSize
        metodo.BarraStato.Panels("utente").Width = metodo.BarraStato.Panels("utente").Width + 20
        metodo.BarraStato.Panels("Designer").Width = 0 'AutoSize
        metodo.BarraStato.Panels("Designer").Width = metodo.BarraStato.Panels("Designer").Width + 35
        
        'Reimposto le immagini a run-time, altrimenti perde il maskcolor (?!?)
        'metodo.BarraStato.Panels("utente").Picture = metodo.ImglistBottoniXP.ListImages("utente").Picture
        'metodo.BarraStato.Panels("dittaanno").Picture = metodo.ImglistBottoniXP.ListImages("dbase").Picture
        'metodo.BarraStato.Panels("Lingua").Picture = metodo.ImglistBottoniXP.ListImages("Euro").Picture
        On Local Error Resume Next
        Dim pct As StdPicture
        If Dir$(MXNU.PercorsoPgm & "\LanguageResources\Flags\Flag" & MXNU.LinguaAttiva & ".ico", vbNormal) <> "" Then
            'metodo.BarraStato.Panels("Lingua").Picture = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\Flag" & MXNU.LinguaAttiva & ".ico")
            
            Set pct = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\Flag" & MXNU.LinguaAttiva & ".ico")
            
            metodo.CommandBars.AddIconHandle pct.Handle, STATUSBAR_ID_PANELLINGUA, 0, False
            metodo.BarraStato.Panels("Lingua").IconIndex = STATUSBAR_ID_PANELLINGUA
        Else
            'metodo.BarraStato.Panels("Lingua").Picture = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\FlagEuro.ico")
            Set pct = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\FlagEuro.ico")
            metodo.CommandBars.AddIconHandle pct.Handle, STATUSBAR_ID_PANELLINGUA, 0, False
            metodo.BarraStato.Panels("Lingua").IconIndex = STATUSBAR_ID_PANELLINGUA
            
        End If
        On Local Error GoTo 0
        'metodo.BarraStato.Panels("Designer").Picture = metodo.ImglistBottoniXP.ListImages("VersDesigner").Picture
    #End If
End Sub

Function ErroreVB(IDErrore%, Msg$) As Integer
    Dim strErrore$, q%
    strErrore = "Errore " & IDErrore & ": "
    ErroreVB = True
    Select Case IDErrore
        Case 5 'Illegal Function Call
            strErrore = strErrore & "Chiamata a funzione non valida."
        Case 6 'OverFlow
            strErrore = strErrore & "Valore fuori limite."
        Case 9 'SubScript Out of Range
            strErrore = strErrore & "Indice di vettore fuori limite."
        Case 11 'Division by Zero
            strErrore = strErrore & "Divisione per zero."
        Case 13 'Type Mismatch
            strErrore = strErrore & "Errore di conversione di dati."
        Case 94 'Invalid use of NULL
            strErrore = strErrore & "Uso non valido dell'operatore NULL."
        Case 380 'Invalid property Value
            strErrore = strErrore & "Settaggio di Proprietà non valido."
        Case ERR_UTENTE
            If Msg = "" Then
                strErrore = "Operazione Fallita."
            Else
                strErrore = Msg
            End If
            ErroreVB = False
        Case Else
            strErrore = "Errore generico: " & Error$
    End Select
    q = MXNU.MsgBoxEX(strErrore, vbExclamation, "ATTENZIONE!")
End Function

Sub Focus_Oggetto(ByVal Desthwnd%, ByVal hwnd&)
    Dim crct As RECT
    Dim rct As RECT
    Dim destDC%, res%

    GetWindowRect Desthwnd%, crct
    GetWindowRect hwnd&, rct
    'differenza
    rct.Left = rct.Left - crct.Left
    rct.Right = rct.Right - crct.Left
    rct.Top = rct.Top - crct.Top
    rct.Bottom = rct.Bottom - crct.Top

    MsgBox "InflateRect rct, 3, 3"

    destDC% = GetDC(Desthwnd%)
    MsgBox "DrawFocusRect destDC, rct"
    res = ReleaseDC(Desthwnd%, destDC)
End Sub


Function FormLoader(FRM As Form, HelpContextID As Long) As Integer
    FormLoader = False
    #If ISMETODO2005 <> 1 Then
        If metodo.mnuMenu.Enabled = False Then
            FormLoader = True
            Exit Function
        End If
    #End If
    On Local Error Resume Next
    FRM.FormProp.FormID = HelpContextID
    On Local Error GoTo out_of_memory
    Load FRM
    Call AssegnaToolTip(FRM)
    
    'FRM.HelpContextID = HelpContextID
        
'    'REMIND: da decommentare a sviluppo terminato
'    '*** modifica ExtensionLoader ***
'    If (Not Designer Is Nothing) Then
'        Dim colObj As Collection
'        Dim colAmb As Collection
'
'        On Local Error Resume Next
'        Set colObj = New Collection
'        colObj.Add hndDBArchivi
'        Set colAmb = Ambienti2Collection(False)
'        Call Designer.LoadExtension(frm, colAmb, colObj)
'    End If
    
    With FRM
        .ZOrder vbBringToFront
        .WindowState = vbNormal
    End With
    On Local Error Resume Next
    Call FRM.MWAgt1.RegistraAgenteFrm(FRM)
    FormLoader = True
    metodo.MousePointer = Default
    On Local Error GoTo 0
fine_Load:
    Exit Function
out_of_memory:
     If Err = 7 Then Resume Next Else Resume fine_Load

End Function

'Sviluppo nr. 1570
Private Sub AssegnaToolTip(FRM As Form)
    Dim objCtrl As Control

    On Local Error Resume Next
    If MXNU.LinguaAttiva <> "IT" Then
        For Each objCtrl In FRM.Controls
            Select Case TypeName(objCtrl)
                Case "Label"
                    objCtrl.ToolTipText = Replace(objCtrl.Caption, "&", "")
                Case "MWEtichetta", "MWLinguetta"
                    objCtrl.Lingua = MXNU.LinguaAttiva
            End Select
        Next
    End If
    On Local Error GoTo 0
    
End Sub

#If ISMETODO2005 <> 1 Then
'Fatta Globale per tutti i moduli
Sub GestioneToolBut(ByVal Button As Object)

    Dim frmAttiva As Form
    Dim bolCancellaAzione As Boolean
    Dim intAzioneEseguita As Integer
    Dim bolApriNomiCtrl As Boolean
    Dim bolAllInOne As Boolean
    Dim DesignButton As MSComctlLib.Button ' DESIGNER
    Dim frmDesign As Form
    
    Dim intForm As Integer
    Dim bolFrmZoomAperta As Boolean
    Dim objExt As Object
                                    
    
    If Button.Enabled Then
        On Local Error Resume Next
        Set frmAttiva = metodo.FormAttiva
        metodo!helpTimer.Interval = 500
        If TypeName(Button) = "IButtonMenu" Then
            ' Gestione Zoom su foglio scheda sviluppo n.ro 1128 + Designer
            If metodo.Barra.Buttons(idxBottoneDefAgenti).Value = tbrUnpressed And LCase$(Left$(Button.key, 4)) <> "zoom" And LCase$(Left$(Button.key, 6)) <> "design" Then
            'If metodo.Barra.Buttons(idxBottoneDefAgenti).Value = tbrUnpressed Then
                Select Case Button.Index
                    Case 1 'Def. Agenti
                        ' Rif. anomalia n.ro 4824 + rif. anomalia 6882
                        If MXNU.ModuloRegole And Not (MXNU.CtrlAccessi) Then
                            Call FormImpostaAgenti(frmAttiva)
                        End If
                    Case 2 'Nome Controlli - campi
                        If (metodo.FormsCount > 2) Then
                            bolApriNomiCtrl = True
                            Call frmAttiva.AzioniMetodo(MetFMostraCampiDBAnagr, bolApriNomiCtrl)
                            If bolApriNomiCtrl Then
                                Set FrmNomiControlli.frmDef = frmAttiva
                                FrmNomiControlli.Show
                            End If
                        End If
                    Case 3 'Situazione Anagrafica
                        'mostro le dipendenze dell'anagrafica
                        Dim ListaCol As New Collection
                        Call frmAttiva.AzioniMetodo(MetFVisDipendenze, ListaCol)
                        If ListaCol.Count > 0 Then
                            Call MXVA.VisualizzaDipendenze(ListaCol)
                        End If
                        Set ListaCol = Nothing
                    Case 4 'Attiva\Disattiva Log
                        If Button.Tag = "A" Then
                            Call MXDB.DisattivaLog
                            Button.Tag = ""
                            Button.Text = "Attiva File Log"
                            Call frmLog.MostraFileLog(MXNU.GetTempDir & "MWSqlLog" & MXNU.NTerminale & ".sql")
                        Else
                            Button.Tag = "A"
                            Call MXDB.AttivaLog
                            Button.Text = "Disattiva File Log"
                        End If
                    Case 5 'ricarica profili
                        metodo.MousePointer = vbHourglass
                        MXNU.MostraMsgInfo 70058
                        '>>> profili anagrafiche
                        If Not MXVA Is Nothing Then
                            If MXVA.ChiudiDyTRAnagraf() Then
                                Call MXVA.ApriDyTRAnagraf
                            End If
                            If MXVA.ChiudiDyTRValidazione() Then
                                Call MXVA.ApriDyTRValidazione
                            End If
                        End If
                        '>>> profili tabelle
                        If Not MXCT Is Nothing Then
                            If MXCT.ChiudiDyTRTabelle() Then
                                Call MXCT.ApriDyTRTabelle
                            End If
                        End If
                        '>>> profili visioni/situazioni
                        If Not MXVI Is Nothing Then
                            If MXVI.ChiudiDyTRSituazioni() Then
                                Call MXVI.ApriDyTRSituazioni
                            End If
                            If MXVI.ChiudiDyTRVisioni() Then
                                Call MXVI.ApriDyTRVisioni
                            End If
                        End If
                        MXNU.MostraMsgInfo ""
                        metodo.MousePointer = vbDefault
                End Select
            ElseIf LCase$(Left$(Button.key, 4)) = "zoom" Then ' MYERP
                 #If ISMETODOXP = 1 Then
                    If MXNU.MetodoXP Then
                        ' Rif. anomalia ISV n.ro 6 / Anomalia Metodo XP n.ro 6084 / Quesito ISV 294
                        Dim objActiveControl As Object
                        If LCase(frmAttiva.Name) = "frmextchild" Then
                            If frmAttiva.ActiveControl.object.Controls.Count > 1 Then
                                Set objActiveControl = frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl
                            Else
                                Set objActiveControl = frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl
                            End If
                        ElseIf TypeName(frmAttiva.ActiveControl) = "MWSchedaBox" Then
                            ' Sono in una scheda
                            If frmAttiva.ActiveControl.ControlsEx.Count > 0 Then
                                ' La scheda contiene delle estensioni
                                ' Setto lo UserControl dentro il wrapper in objExt (ExtWrapper deve avere una public property get come tutti gli UserControl estensioni)
                                'Rif. anomalia #8088
                                Set objExt = frmAttiva.ActiveControl.ControlsEx(0).object.Controls
                                If TypeName(objExt(0)) = "CTLXBus" Then
                                    Set objActiveControl = objExt(1).object.ExtActiveControl
                                Else
                                    Set objActiveControl = objExt(0).object.ExtActiveControl
                                End If
                                Set objExt = Nothing
                            End If
                        Else
                            Set objActiveControl = frmAttiva.ActiveControl
                        End If
                        If TypeName(objActiveControl) = "fpSpread" Then
                            ' Gestione dell'apertura di una singola form di zoom
                            bolFrmZoomAperta = False
                            intForm = 0
                            While (intForm < VB.Forms.Count) And (Not (bolFrmZoomAperta))
                                If UCase(VB.Forms(intForm).Name) = "FRMZOOM" Then
                                    bolFrmZoomAperta = True
                                End If
                                intForm = intForm + 1
                            Wend
                            If Not bolFrmZoomAperta Then
                                Call frmZoom.DoZoom(frmAttiva, objActiveControl, Button.Index, objActiveControl.ActiveCol, objActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
                            Else
                                Call MXNU.MsgBoxEX(2738, vbExclamation, 23481)
                            End If
                        End If
                        Set objActiveControl = Nothing
                    End If
                 #End If
            ElseIf LCase$(Left$(Button.key, 6)) = "design" Then ' DESIGNER
                #If ISMETODOXP = 1 Then
                  If MXNU.MetodoXP Then
                     Select Case LCase(Button.key)
                       Case "design", "designer"
                         If Not (frmAttiva Is Nothing) Then
                            ' Rif. anomalia n.ro 6082
                            If frmAttiva.Name <> "frmVisioni" And frmAttiva.Name <> "frmTabelle" And _
                            frmAttiva.Name <> "frmModuli98" And frmAttiva.Name <> "metodo" And _
                             frmAttiva.Name <> "frmToolTip" And frmAttiva.Name <> "frmModuliXP" And frmAttiva.Name <> "frmZoom" Then
                           'If frmAttiva.Name <> "frmExtChild" And frmAttiva.Name <> "frmTabelle" And frmAttiva.Name <> "frmModuli98" And frmAttiva.Name <> "metodo" And _
                             frmAttiva.Name <> "frmToolTip" And frmAttiva.Name <> "frmModuliXP" Then
                                If metodo!tlbDesigner.Visible = False Then
                                    'Se premuto visualizza la barra del designer (l'utente è sicuramente amministratore)
                                    metodo!cmbDesignType.AddItem MXNU.CaricaStringaRes(75474)
                                    metodo!cmbDesignType.AddItem MXNU.CaricaStringaRes(75475)
                                    metodo!tlbDesigner.Visible = True
                                    For Each DesignButton In metodo!tlbDesigner.Buttons
                                        DesignButton.Enabled = False
                                    Next
                                    
                                    'Rimane sempre attivo il bottone "showdiff"
                                    metodo!tlbDesigner.Buttons.Item("showdiff").Enabled = True
                                    ' "Spengo" i bottoni ...
                                    metodo!tlbDesigner.Buttons.Item("showdiff").Value = tbrUnpressed
                                    metodo!tlbDesigner.Buttons.Item("unvisible").Value = tbrUnpressed
                                    metodo!tlbDesigner.Buttons.Item("disable").Value = tbrUnpressed
                                    metodo!cmbDesignType.Enabled = True
                                    ' Quando attivo la barra dovrò al primo salvataggio scegliere la versione
                                    Designer.bolSalvato = False
                                ElseIf metodo!tlbDesigner.Visible = True Then
                                    ' Nascondo la barra del designer
                                    If Designer.bolModify Then
                                        If MXNU.MsgBoxEX(2631, vbYesNo + vbDefaultButton2, 2632) = vbYes Then
                                            Call Designer.TerminateEditMode
                                            Call Designer.ExportDesign(False)
                                        Else
                                            Call Designer.TerminateEditMode
                                        End If
                                        Call Designer.TerminateDesigner
                                        Designer.bolModify = False
                                    End If
                                    metodo!cmbDesignType.Clear
                                    metodo!tlbDesigner.Visible = False
                                    
                                    For Each frmDesign In VB.Forms
                                        If LCase(frmDesign.Name) = "frmdesignerprop" Then
                                            Unload (FrmDesignerProp)
                                        End If
                                    Next
                                End If
                           End If
                         End If
                       Case "designdelete"
                         ' Cancellazione delle versioni
                         Load frmDeleteVersions
                         frmDeleteVersions.Show
                       Case "designversions"
                         Load frmUsersVersions
                         frmUsersVersions.Show
                       Case "designgrid"
                         Load FrmDesignerGrid
                         FrmDesignerGrid.Show
                     End Select
                  End If
                #End If
            End If
        Else
            Err.Clear
            Select Case Button.Index
                Case idxBottoneModuli
                    #If ISMETODOXP = 1 Then
                        If MXNU.MetodoXP Then

                            If frmModuli.StatoMenuAlbero(0) = False Then
                                'il menu moduli non e' aperto ne lockato quindi lo apro e lo locko
                                frmModuli.StatoMenuModuli = MnuModATTIVO
                                Call frmModuli.AlberoBloccatoMenu(True, 0)
                            Else
                                'il menu moduli e' gia' aperto e lockato, lo chiudo
                                'frmModuli.StatoMenuModuli = MnuModNONATTIVO
                                Call frmModuli.AlberoBloccatoMenu(False, 0)
                            End If
                        Else
                            Call AttivaMenuMetodo
                        End If
                    #Else
                        Call AttivaMenuMetodo
                    #End If
                
                Case idxBottoneAllInOne
                    'If MXNU.MetodoXP Then
                        bolAllInOne = True
                        Call frmAttiva.AzioniMetodo(MetFAllInOne, bolAllInOne)
                        If bolAllInOne Then
                            Call AllInOneManager(MXNU.FrmMetodo.FormAttiva)
                        End If
                    'End If
                    
                Case idxBottoneInserisci
                    Call EseguiAgenteAzioneForm(frmAttiva, "before_insert", bolCancellaAzione)
                    If Not bolCancellaAzione Then
                        Call DoLostFocus(frmAttiva.ActiveControl)
                        intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFInserisci))
                        'call frmAttiva.AzioniMetodo(MetFInserisci)
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_insert", bolCancellaAzione, intAzioneEseguita)
                    End If
    
                Case idxBottoneDettaglio
                    Call DoLostFocus(frmAttiva.ActiveControl)
                    Call frmAttiva.AzioniMetodo(MetFDettagli)
                    
                Case idxBottoneRegistra
                    Call EseguiAgenteAzioneForm(frmAttiva, "before_save", bolCancellaAzione)
                    If Not bolCancellaAzione Then
                        Call DoLostFocus(frmAttiva.ActiveControl)
                        intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFRegistra))
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_save", bolCancellaAzione, intAzioneEseguita)
                    End If
                                
                Case idxBottoneAnnulla
                    Call EseguiAgenteAzioneForm(frmAttiva, "before_delete", bolCancellaAzione)
                    If Not bolCancellaAzione Then
                        Call DoLostFocus(frmAttiva.ActiveControl)
                        intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFAnnulla))
                        'Call frmAttiva.AzioniMetodo(MetFAnnulla)
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_delete", bolCancellaAzione, intAzioneEseguita)
                    End If
                
                Case idxBottonePrimo
                    Call EseguiAgenteAzioneForm(frmAttiva, "before_first", bolCancellaAzione)
                    If Not bolCancellaAzione Then
                        Call DoLostFocus(frmAttiva.ActiveControl)
                        Call frmAttiva.AzioniMetodo(MetFPrimo)
                    End If
                    Call EseguiAgenteAzioneForm(frmAttiva, "after_first", bolCancellaAzione)
                
                Case idxBottonePrecedente
                    Call EseguiAgenteAzioneForm(frmAttiva, "before_previous", bolCancellaAzione)
                    If Not bolCancellaAzione Then
                        Call DoLostFocus(frmAttiva.ActiveControl)
                        Call frmAttiva.AzioniMetodo(MetFPrecedente)
                    End If
                    Call EseguiAgenteAzioneForm(frmAttiva, "after_previous", bolCancellaAzione)
                
                Case idxBottoneSuccessivo
                    Call EseguiAgenteAzioneForm(frmAttiva, "before_next", bolCancellaAzione)
                    If Not bolCancellaAzione Then
                        Call DoLostFocus(frmAttiva.ActiveControl)
                        Call frmAttiva.AzioniMetodo(MetFSuccessivo)
                    End If
                    Call EseguiAgenteAzioneForm(frmAttiva, "after_next", bolCancellaAzione)
                Case idxBottoneUltimo
                    Call EseguiAgenteAzioneForm(frmAttiva, "before_last", bolCancellaAzione)
                    If Not bolCancellaAzione Then
                        Call DoLostFocus(frmAttiva.ActiveControl)
                        Call frmAttiva.AzioniMetodo(MetFUltimo)
                    End If
                    Call EseguiAgenteAzioneForm(frmAttiva, "after_last", bolCancellaAzione)
                Case idxBottoneStampa
                    Call frmAttiva.AzioniMetodo(MetFStampa)
                Case idxBottoneTrova
                    If Not MXCT.TTrova(FrmTrovaGen) Then
                        Call frmAttiva.AzioniMetodo(MetFTrova)
                    End If
                Case idxBottoneDefAccessi
                    Call FormDefinisciAccessi(frmAttiva)
                Case idxBottoneHelp
                    #If ISMETODOXP = 1 Then
                        If MXNU.MetodoXP = True Then
                                #If TOOLS = 1 Then
                                    If MXNU.DammiFormAttiva.Name = "frmModuli" Then
                                        Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
                                    Else
                                        Call MXNU.ApriHelp(False)
                                    End If
                                #Else
                                    'posso caricare l'help direttamente su il browser della form tooltip
                                    If MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "METODOW", "HelpCompilato", 0) = 1 Then
                                        Call MXNU.ApriHelp(False)
                                    End If
                                #End If
                        Else
                            #If TOOLS = 1 Then
                                If MXNU.DammiFormAttiva.Name = "frmModuli" Then
                                    Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
                                Else
                                    Call MXNU.ApriHelp(False)
                                End If
                            #Else
                                Call MXNU.ApriHelp(False)
                            #End If
                        End If
                    #Else
                        #If TOOLS = 1 Then
                            If MXNU.DammiFormAttiva.Name = "frmModuli" Then
                                Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
                            Else
                                Call MXNU.ApriHelp(False)
                            End If
                        #Else
                            Call MXNU.ApriHelp(False)
                        #End If
                    #End If
                Case idxBottoneUtenteModifica
                    Call frmAttiva.AzioniMetodo(MetFVisUtenteModifica)
                Case idxBottoneDefAgenti 'Def Agenti predefiniti
                    If (metodo.pBottoniAgenti = ageCollegamenti) Then
                        'imposto agenti nella form
                        If MXNU.ModuloRegole And Not (MXNU.CtrlAccessi) Then   ' rif. anomalia n.ro 4824+ n.ro 6882
                            Call FormImpostaAgenti(frmAttiva)
                        End If
                    End If

                Case idxBottoneSchedula
                    Call frmAttiva.AzioniMetodo(MetFSchedulaOperazione)
                Case idxBottoneSpostaBarra
                    metodo!Barra.Align = (metodo!Barra.Align) Mod 2 + 1
                    metodo!Barra.Buttons(Button.Index).Image = metodo!Barra.Align + 10
                    metodo!tlbDesigner.Align = (metodo!tlbDesigner.Align) Mod 2 + 1 ' Designer
                Case idxBottoneZoom
                    #If ISMETODOXP = 1 Then
                      If MXNU.MetodoXP Then
                        Dim intZoomType As Integer
                        intZoomType = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\Mw.ini", MXNU.UtenteAttivo, "ZoomType", 1)
                        ' Gestione dell'apertura di una singola form di zoom
                        bolFrmZoomAperta = False
                        intForm = 0
                        While (intForm < VB.Forms.Count) And (Not (bolFrmZoomAperta))
                            If UCase(VB.Forms(intForm).Name) = "FRMZOOM" Then
                                bolFrmZoomAperta = True
                            End If
                            intForm = intForm + 1
                        Wend
                        
                        If Not bolFrmZoomAperta Then
                            ' Rif. anomalia ISV n.ro 6 / Anomalia Metodo XP n.ro 6084 / Quesito ISV 294
                            If LCase(frmAttiva.Name) = "frmextchild" Then
                                If frmAttiva.ActiveControl.object.Controls.Count > 1 Then
                                    Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl, intZoomType, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl.ActiveCol, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl.ActiveRow, MXNU.FrmMetodo!Barra.Buttons(idxBottoneTrova).Enabled)
                                Else
                                    Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl, intZoomType, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl.ActiveCol, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl.ActiveRow, MXNU.FrmMetodo!Barra.Buttons(idxBottoneTrova).Enabled)
                                End If
                            ElseIf TypeName(frmAttiva.ActiveControl) = "MWSchedaBox" Then
                                ' Sono in una scheda
                                If frmAttiva.ActiveControl.ControlsEx.Count > 0 Then
                                    ' La scheda contiene delle estensioni
                                    ' Setto lo UserControl dentro il wrapper in objExt (ExtWrapper deve avere una public property get come tutti gli UserControl estensioni)
                                    Set objExt = frmAttiva.ActiveControl.ControlsEx(0).object.Controls
                                    'Rif anomalia #8088
                                    If TypeName(objExt(0)) = "CTLXBus" Then
                                        Set objActiveControl = objExt(1).object.ExtActiveControl
                                    Else
                                        Set objActiveControl = objExt(0).object.ExtActiveControl
                                    End If
                                    Set objExt = Nothing
                                    Call frmZoom.DoZoom(frmAttiva, objActiveControl, intZoomType, objActiveControl.ActiveCol, objActiveControl.ActiveRow, MXNU.FrmMetodo!Barra.Buttons(idxBottoneTrova).Enabled)
                                    Set objActiveControl = Nothing
                                End If
                            Else
                                Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl, intZoomType, frmAttiva.ActiveControl.ActiveCol, frmAttiva.ActiveControl.ActiveRow, MXNU.FrmMetodo!Barra.Buttons(idxBottoneTrova).Enabled)
                            End If
                        Else
                            Call MXNU.MsgBoxEX(2738, vbExclamation, 23481)
                        End If
                        
                      End If
                    #End If
                Case idxBottoneDesigner
                  #If ISMETODOXP = 1 Then
                    If MXNU.MetodoXP Then
                      If Not (frmAttiva Is Nothing) Then
                        ' Rif. anomalia n.ro 6082
                        If frmAttiva.Name <> "frmVisioni" And frmAttiva.Name <> "frmTabelle" And frmAttiva.Name <> "frmModuli98" And _
                            frmAttiva.Name <> "metodo" And frmAttiva.Name <> "frmZoom" And _
                            frmAttiva.Name <> "frmToolTip" And frmAttiva.Name <> "frmMetodoXP" And frmAttiva.Name <> "frmModuliXP" Then
                            If metodo!tlbDesigner.Visible = False Then
                                'Se premuto visualizza la barra del designer (utente amministratore)
                                metodo!cmbDesignType.AddItem MXNU.CaricaStringaRes(75474)
                                metodo!cmbDesignType.AddItem MXNU.CaricaStringaRes(75475)
                                metodo!tlbDesigner.Visible = True
                                For Each DesignButton In metodo!tlbDesigner.Buttons
                                    DesignButton.Enabled = False
                                Next
                                'Rimane sempre attivo il bottone "showdiff" (e non premuto)
                                metodo!tlbDesigner.Buttons.Item("showdiff").Enabled = True
                                metodo!tlbDesigner.Buttons.Item("showdiff").Value = tbrUnpressed
                                 
                                metodo!cmbDesignType.Enabled = True
                                ' Quando attivo la barra dovrò al primo salvataggio scegliere la versione
                                Designer.bolSalvato = False
                            ElseIf metodo!tlbDesigner.Visible = True Then
                                'Nascondo la barra del designer
                                If Designer.bolModify Then
                                    If MXNU.MsgBoxEX(2631, vbYesNo + vbDefaultButton2, 2632) = vbYes Then
                                        Call Designer.TerminateEditMode
                                        Call Designer.ExportDesign(False)
                                    Else
                                        Call Designer.TerminateEditMode
                                    End If
                                    Call Designer.TerminateDesigner
                                    Designer.bolModify = False
                                End If
                                metodo!cmbDesignType.Clear
                                metodo!tlbDesigner.Visible = False
                                For Each frmDesign In VB.Forms
                                    If LCase(frmDesign.Name) = "frmdesignerprop" Then
                                        Unload (FrmDesignerProp)
                                    End If
                                Next
                            End If
                        End If
                      End If
                    End If
                  #End If
               Case 30
                  Call ShellExecute(metodo.hwnd, "Open", "www.metodo.it/MetodoPortal/DesktopPortalMain.aspx?action=NEWS", "", App.Path, 1)
            End Select
        End If
        Set frmAttiva = Nothing
    End If
    
        
End Sub
#End If

Function GetUltimoAnnoUsato() As String
    Dim riga$
    riga$ = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "TRM" & MXNU.NTerminale, MXNU.DittaAttiva, MXNU.AnnoDflt)
    If riga$ <> "" Then
    GetUltimoAnnoUsato$ = riga
    Else
    GetUltimoAnnoUsato$ = ""
    End If
    'Dim hSS%, q%, sql$
    'sql = "SELECT * FROM TabUsoDitte WHERE ((NumTerminale =" & Val(NTerminale$) & ") AND (Ditta ='" & MXNU.DittaAttiva & "'))"
    'hSS = MXDB.dbCreaSS(hndDBDitte, sql, TIPO_TABELLA)
    'If MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
    '   GetUltimoAnnoUsato$ = ""
    'Else
    '    GetUltimoAnnoUsato$ = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Anno", "")
    'End If
    'q = MXDB.dbChiudiSS(hSS)
End Function
Function LeggiDscBreve(strDitta As String) As String
    LeggiDscBreve = LeggiDscBreve_(strDitta, "")
End Function

Function LeggiDscBreve_(Ditta As String, dflt As String) As String
Dim intq As Integer
Dim hSS As MXKit.CRecordSet
Dim strSQL As String

'    Select Case intMetExe
'        Case EXE_METODO95
            strSQL = "SELECT DataCostituzione,DesBreve FROM  TabDitte"
'        Case EXE_RITENUTE, EXE_CESPITI
'            sqlt = "SELECT DesBreve FROM  AnagraficaDitte WHERE Ditta ='" + Ditta + "'"
'    End Select
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    intq = Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    If intq Then
        LeggiDscBreve_ = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DesBreve", dflt)
        MXNU.DataCostituzione = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DataCostituzione", MXNU.Default_Data)
    Else
        LeggiDscBreve_ = dflt
    End If
    intq = MXDB.dbChiudiSS(hSS)
End Function

Sub Main()
    Dim intRis As Boolean
    Dim strUtente As String
    Dim bolConfInt As Boolean
    Dim strLineErr As String
    Dim strDateLayout As String
    
    On Local Error GoTo err_Main

    InitCommonControls   'Necessario per il corretto rendering degli oggetti in 3D col Manifest
    
    strLineErr = "Creazione Oggetto Nucleo"
    Set MXNU = New MXNucleo.XNucleo

    Load frmIntro
    '>>> INIZIALIZZAZIONE NUCLEO
    #If TOOLS = 1 Then
        MXNU.FileDatLocali = False
    #End If
    strLineErr = "Inizializzazione Oggetto Nucleo"
    
    strinitexe = Command$
    If (Not MXNU.Inizializza(Command$, strUtente)) Then GoSub err_objInit
    
    ' Inizializzazione del charset per la gestione delle lingue (polacco, slovacco, ecc.)
    Call InitCharSet(MXNU.PercorsoPgm & "\" & "CharSetManager.config")
    
    ' Spostato i controlli su carattere separatore dei decimali e del controllo della data
    '<rif anomalia 816 RZ>
    If QueryValue(HKEY_CURRENT_USER, "Control Panel\International", "sMonDecimalSep") <> "," Then 'And MXNU.LinguaAttiva = "IT"
        Call MXNU.MsgBoxEX("Impostare il carattere ',' (Virgola) nelle Impostazioni Internazionali - Valuta - Separatore Decimali da Pannello di Controllo !", vbOKOnly + vbCritical, 1007)
        GoTo err_objInit
    End If

    strDateLayout = QueryValue(HKEY_CURRENT_USER, "Control Panel\International", "sShortDate")
    If Len(strDateLayout) < 10 Then
        Call MXNU.MsgBoxEX(3065, vbOKOnly + vbCritical, 1007, strDateLayout)
        GoTo err_objInit
    End If
    
    '<rif anomalia 9303 RZ>
    If Mid(strDateLayout, 3, 1) <> "/" Then
        Call MXNU.MsgBoxEX("Impostare la data nel formato dd/MM/yyyy", vbOKOnly + vbCritical, 1007, strDateLayout)
        GoTo err_objInit
    End If
    
    
    MXNU.VersioneMetodo = App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
    MXNU.EXEName = App.EXEName
    
    '[16/06/2011] Rimozione Chiave Hardware
'    '>>> RICERCA CHIAVE HARDWARE
'    strLineErr = "Ricerca Chiave Hardware"
'    frmIntro.MostraMessaggioOperazione (9008)
    
'>>>>>>>>>>> SPOSTATO SU MDIFORM_LOAD, ALTRIMENTI I COLORI DEL GRADIENTE NON SI INIZIALIZZANO CORRETTAMENTE <<<<<<<<<<<<<<<<<<<<<
'    'imposto la variabile di controllo ISMETODO2005 sui controlli
'    strLineErr = "Inizializzazione sistema Metodo 2005"
'    Dim oM2005 As MXCtrl.M2005Setup
'    Set oM2005 = New MXCtrl.M2005Setup
'#If ISMETODO2005 = 1 Then
'    oM2005.ISMETODO2005 = True
'#Else
'    oM2005.ISMETODO2005 = False
'#End If
'    Set oM2005 = Nothing
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    strLineErr = "Inizializzazione Oggetto Spread"
'>>>>>>>>>>> SPOSTATO SU MDIFORM_LOAD, VEDI SOPRA <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'#If ISMETODO2005 = 1 Then
'    Call InizializzaSpread(MXNU.MetodoXP, True, SysGradientColor1)
'#Else
    Call InizializzaSpread(MXNU.MetodoXP)
'#End If
    
#If ISKEY Then
    funzione1
#End If
    strLineErr = "Creazione Oggetti KIT - BUSINESSS"
    
    '>>> CREAZIONE OGGETTI KIT - BUSINESS
#If SOLOKIT = 1 Then
    If Not CreateObjKitBus(frmIntro.CTLXKit1, Nothing) Then GoSub err_objInit
#Else
    If Not CreateObjKitBus(frmIntro.CTLXKit1, frmIntro.CTLXBus1) Then GoSub err_objInit
#End If
    strLineErr = "Inizializzazione Libreria Database"
    
    '>>> INIZIALIZZAZIONE LIBRERIA DATABASE
    If Not (MXDB.dbInizializza(MXNU)) Then
        Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Libreria ODBC"))
        GoSub err_objInit
    End If

    strLineErr = "Apertura Ditta"
    
    '>>> APERTURA DITTA
    #If ISKEY <> 1 Then
        If Not ApriDitta(MXNU.UtenteDB, MXNU.PasswordDB) Then
            Call DropObjKitBus
            GoSub err_objInit
        End If
    #End If
    
    strLineErr = "Inizializzazione Oggetti KIT - BUSINESS"
    '>>> INIZIALIZZAZIONE OGGETTI KIT - BUSINESS
    If Not InitObjKitBus(hndDBArchivi) Then
        Call DropObjKitBus
        GoSub err_objInit
    End If

    strLineErr = "Apertura Anno"
    '>>> SELEZIONE ANNO
    #If ISKEY <> 1 Then
        Dim NuovoAnno As Integer
        If SelezioneAnno(False, NuovoAnno) Then
            MXNU.AnnoAttivo = NuovoAnno
            Call ApriAnno(False)
        Else
            Call DropObjKitBus
            GoSub err_objInit
        End If
    #Else

    #End If

    strLineErr = "Connessione MetServer"
    #If USAM98SERVER Then
    Call ConnectM98Server(GobjM98Server)
    #End If

    strLineErr = "Supporto Script"
    'supporto Script
    AddAmbienti2Script
    
    strLineErr = "Inizializzazione Designer"
    '*** DESIGNER ***
    ' controlla il flag in MW.INI
    ' inizio rif.sch. A5714
    Dim strEntryMw As String
    Dim bolDesignerAbilitato As Boolean
    bolDesignerAbilitato = True
    strEntryMw = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MW.INI", "METODOW", "ATTIVADESIGNER", "")
    If (Len(strEntryMw) > 0) Then
        bolDesignerAbilitato = (StrComp(strEntryMw, "DESIGNER", vbTextCompare) = 0)
    End If
    If bolDesignerAbilitato Then
        'Controllo esistenza modulo runtime nella chiave
        If ((MXNU.ControlloModuliChiave(modMyErpRunTime) = 0) _
            Or (MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0)) Then 'modMyErpRunTime = 118
            
            Set Designer = New MXDesigner.cDesigner
            Set Designer.MyMXNU = MXNU
            If Not (Designer.Attiva) Then
                Set Designer = Nothing
            End If
        Else
            Set Designer = Nothing
        End If
    Else
        Set Designer = Nothing
    End If
    ' fine rif.sch. A5714
    
    '>>>inizializzazione sistema di messaggistica
#If ISMETODO2005 = 1 Then
    On Local Error GoTo err_Main_NoBlock
    
    'inizializzazione hosting metodo
    strLineErr = "Inizializzazione Metodo Interop"
    Set mMetodoInterop = New CMetodoInterop
    
    'innizializzazione browser
    strLineErr = "Inizializzazione Metodo Browser"
    Set mMetodoBrowser = CreateObject("MxBrowser.CBrowserEngine")
    If (Not mMetodoBrowser Is Nothing) Then
            If (Not mMetodoBrowser.Initialize(mMetodoInterop, metodo.DockingPaneManager)) Then
                    Set mMetodoBrowser = Nothing
            End If
    End If
    
    'gestione messaggistica da file ini <rif anomalia #8620 RZ>
    Dim msgingInstalled As String
    msgingInstalled = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", MXNU.UtenteSistema, "Messaggistica", "")
    If msgingInstalled = Empty Then
        msgingInstalled = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "METODOW", "Messaggistica", 1)
    End If
    
    If msgingInstalled > 0 Then
        strLineErr = "Inizializzazione sistema di messaggistica"
        On Local Error Resume Next
        Set mMessagingEngine = CreateObject("MxMailing.MailingEngine")
        Dim myerr As String
        Dim myerrdescription As String
        Dim myerrconfig As String
        myerr = Err.Number
        myerrdescription = Err.Description
        On Local Error GoTo 0
        If mMessagingEngine Is Nothing And Err = 429 Then 'il sistema di messaggistica non è correttamente registrato
            Call MsgBox("Errore [" & myerr & " " & myerrdescription & "] nella funzione [main] durante l'operazione " & strLineErr, vbCritical, "Attenzione!")
        Else
            If (mMessagingEngine.Initialize(MXNU, MXNU.UtenteSistema, MXNU.PercorsoPgm, myerrconfig)) Then
                Call mMessagingEngine.InitializeActionExecutor(Ambienti2Collection(False), hndDBArchivi)
            End If
            If myerrconfig <> Empty Then
                Call MXNU.MsgBoxEX(myerrconfig, vbCritical, "Metodo Evolus Messaging")
            End If
        End If
    End If
    
    On Local Error GoTo err_Main
#End If
    
    strLineErr = "Load Metodo"
    '>>> APERTURA METODO
    Load metodo
    
    #If ISMETODOXP = 1 Then
        If MXNU.MetodoXP Then
            metodo.BarraStato.Panels("Designer").Text = MXNU.VersioneAttiva
        End If
    #End If
    
    'Sviluppo 2033: Agente Fisso in partenza di Metodo
    If MXAA.AgentiFissi Then
        Dim objAgt As MXKit.CAgenteAuto
        Set objAgt = MXAA.CreaCAgenteAuto()
        
        Call MXAA.EseguiAgt(objAgt, "FISSI\STARTUPMETODO")
        
        Set objAgt = Nothing
    End If
    
    'Anomalia 8516 - Spostato chiamata ad AttivaMenu dopo il load della MDI
    #If ISMETODO2005 = 1 Then
        frmModuli.ModuloAttivo = "*DA_INI*"    'Leggo il modulo attivo dall'mw.ini. Vedi funzione AttivaMenu su MetodoXP
    #Else
        frmModuli.AttivaMenu
    #End If
    
fine_Main:
    On Local Error GoTo 0
    Exit Sub

err_Main:
    
    
    If Not (MXNU Is Nothing) Then
        Call MsgBox("Errore [" & Err.Number & " " & Err.Description & "] nella funzione [main] durante l'operazione " & strLineErr, vbCritical, "Attenzione!")
        Set MXNU = Nothing
    Else
        Call MsgBox("Errore [" & Err.Number & " " & Err.Description & "] nella funzione [main]", vbCritical, "Attenzione!")
    End If
    On Local Error Resume Next
    Unload frmIntro
    Resume fine_Main
    Resume
err_objInit:
    If Not (MXNU Is Nothing) Then Set MXNU = Nothing
    On Local Error Resume Next
    Unload frmIntro
    End

err_Main_NoBlock:
    If Not (MXNU Is Nothing) Then
        Call MsgBox("Errore [" & Err.Number & " " & Err.Description & "] nella funzione [main] durante l'operazione " & strLineErr, vbCritical, "Attenzione!")
        'Set MXNU = Nothing   'Remmato altrimenti và in errore successivamente bloccando l'esecuzione di Metodo
    Else
        Call MsgBox("Errore [" & Err.Number & " " & Err.Description & "] nella funzione [main]", vbCritical, "Attenzione!")
    End If
    Resume Next

End Sub

'(rif 10)
Function FormaNumeroDoc$(ByVal anno%, ByVal NrDoc&, bis$)
    FormaNumeroDoc$ = Format$(anno, "0000") & "/" & Format$(CStr(NrDoc&), "00000000") & "/" & UCase$(bis$)
End Function

Sub ScomponiNumeroDoc(AnnoNrBis$, anno%, NrDoc&, bis$)
    anno% = Val(Left$(AnnoNrBis$, 4))
    NrDoc& = Val(Mid(AnnoNrBis$, 6, 8))
    bis$ = Right$(AnnoNrBis$, 1)
End Sub


'Restituisce le quantità LIFO dell'Articolo passato
Function LeggiGiacenzaInizialeLIFO(CodArt As String) As Currency
Dim q As Integer
Dim Sql As String
Dim hndSS As CRecordSet
Dim GILIFO As Currency

    Sql = "SELECT Quantita1,Quantita2,Quantita3,Quantita4,Quantita5,Quantita6,Quantita7,Quantita8,Quantita9,Quantita10,Quantita11,Quantita12,Quantita13,Quantita14,Quantita15 FROM LIFOArticoli WHERE (CodiceArt = '" & CodArt$ & "')"
    Set hndSS = MXDB.dbCreaSS(hndDBArchivi, Sql$, TIPO_TABELLA)
    If MXDB.dbFineTab(hndSS, TIPO_SNAPSHOT) Then
        GILIFO@ = 0
    Else
        For q% = 1 To 15
            GILIFO@ = GILIFO@ + MXDB.dbGetCampo(hndSS, TIPO_SNAPSHOT, "Quantita" & q%, 0@)
        Next q%
    End If
    q% = MXDB.dbChiudiSS(hndSS)
    LeggiGiacenzaInizialeLIFO = GILIFO@
End Function

'Calcola il Valore LIFO TOTALE dell'Articolo Passato
'PARAMETRI:
'   CodArt$     Codice Articolo
'   Qta@        Quantità Carico - Quantità Scarico dell'Articolo
'   PMA@        Valore Medio (=Valore Carico/Qtà Carico)
'   LIFOTot@    Valore LIFO Totale
'   LIFOUn@     Valore LIFO Unitario
Sub ValorizzazioneLIFO(CodArt$, ByVal qta@, ByVal PMA@, LIFOTot@, LIFOUn@)
    Dim hSS As CRecordSet, q%, Sql$, qt@
    Sql$ = "SELECT * FROM LIFOArticoli WHERE (CodiceArt = '" & CodArt$ & "')"
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, Sql, TIPO_TABELLA)
    LIFOTot@ = 0@
    If MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
        LIFOTot@ = qta@ * PMA@
    Else
        Dim i%, qtl@, vl@, qtLifo@
        qtLifo@ = 0
        For i% = 1 To 15
            qtLifo@ = qtLifo@ + MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Quantita" & i, 0@)
        Next i%
        qt@ = qta@ + qtLifo@

        For i = 15 To 1 Step -1
            qtl@ = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Quantita" & i, 0@)
            vl@ = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "ValoreLIFO" & i, 0@)
            If qtl < qt Then
                LIFOTot@ = LIFOTot@ + qtl * vl@
                qt@ = qt@ - qtl@
            Else
                LIFOTot@ = LIFOTot@ + (qt@ * vl@)
                qt@ = 0@
            Exit For
            End If
        Next i
        If qt@ > 0@ Then LIFOTot@ = LIFOTot@ + (qt@ * PMA@)
    End If
    q% = MXDB.dbChiudiSS(hSS)
    qtLifo@ = (qta@ + qtLifo@)
    If qtLifo@ = 0 Then
        LIFOUn@ = 0
    Else
        LIFOUn@ = LIFOTot@ / qtLifo@
    End If
End Sub

Function CaricaDescrLingue(Foglio As Object, lngPrimaColonna As Long)
    Dim lngCodice As Long, hSSLingue As CRecordSet, varDescrizione As Variant
    Dim intq As Integer, strSQL As String

    strSQL = "SELECT Codice,Descrizione FROM TabLingue WHERE Codice > 0"
    Set hSSLingue = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    intq = Not MXDB.dbFineTab(hSSLingue, TIPO_SNAPSHOT)
    Do While intq
        lngCodice = MXDB.dbGetCampo(hSSLingue, TIPO_SNAPSHOT, "Codice", 0)
        varDescrizione = MXDB.dbGetCampo(hSSLingue, TIPO_SNAPSHOT, "Descrizione", 0)
        intq = MXDB.dbSuccessivo(hSSLingue, TIPO_SNAPSHOT)

        Call ssButtonSetPicture(Foglio, lngCodice + lngPrimaColonna, 0, Nothing, varDescrizione)
    Loop
    intq = MXDB.dbChiudiSS(hSSLingue)

End Function

'Stampa il contenuto di un foglio su stampante.
'Parametri:
'   foglio = foglio da stampare
'   PrtColHeaders = True per stampare l'intestazione di colonna,False altrimenti.
'   PrtRowHeaders = True per stampare l'intestazione di riga,False altrimenti.
'   TitoloStampa = Stringa da usare come titolo per la coda di stampa.
'   TitoloTesta = Stringa da usare come testo da stampare nell'intestazione della pagina.
Sub StampaFoglio(Foglio As FPSpreadADO.fpSpread, ByVal PrtColHeaders%, ByVal PrtRowHeaders%, ByVal TitoloStampa$, TitoloTesta$)
    Dim objCRW As MXKit.CCrw
    Set objCRW = MXCREP.CreaCCrw()
    objCRW.ClearOpzioniStp
    objCRW.OpzioniForm = 0&
    objCRW.MostraFrmStampa
    If Not objCRW.Stampa_Annullata Then
        If objCRW.periferica = "Stampante" Then

            'Foglio.hDCPrinter = objCRW.Stampante.hdc
            Foglio.hDCPrinter = 0  'Rif. Anomalie98 Nr. 2267: altrimenti non stampa niente; rimane un solo problema: si stampa solo sulla stampante di default
            If objCRW.Stampante.nMinPage <> 0 And objCRW.Stampante.nMaxPage <> 0 Then
                Foglio.PrintPageStart = objCRW.Stampante.nMinPage
                Foglio.PrintPageEnd = objCRW.Stampante.nMaxPage
                Foglio.PrintType = PrintTypePageRange
            Else
                Foglio.PrintType = PrintTypeAll
            End If
            Foglio.PrintBorder = True
            Foglio.PrintColHeaders = PrtColHeaders
            Foglio.PrintRowHeaders = PrtRowHeaders
            Foglio.PrintColor = True
            Foglio.PrintHeader = "/fn""Arial"" /fz""10"" /fb1   " & Trim$(MXNU.Dsc_Breve_Ditta) & " - " & MXNU.AnnoAttivo & "        " & TitoloTesta & "  /n /fb0 " & Now
            Foglio.PrintFooter = "/fn""Arial"" /fz""10"" /fb0 /r Pag. /p"
            Foglio.PrintGrid = False
            Foglio.PrintJobName = TitoloStampa
            Foglio.PrintShadows = True
            Foglio.PrintUseDataMax = True

            metodo.MousePointer = vbHourglass
            Foglio.Action = ActionPrint
            metodo.MousePointer = vbDefault
        End If
    End If
    Set objCRW = Nothing
End Sub



Sub CaricaTabella(strNomeDbTabella As String, vntDesTabella As Variant, lngHelp As Long, Optional ChiaveAgg As Variant, Optional WHEAgg As Variant)
    Dim intq As Boolean, FrmTab As New frmTabelle, lngHwndForm As Long, ints As Integer
    Dim strChiaveAgg As String
    Dim StrWheAgg As String
    Dim lngDesTabella As Long ' per compatibilità

    FrmTab.NOMETABELLA = strNomeDbTabella
    If IsNumeric(vntDesTabella) Then
        vntDesTabella = "{" & Val(vntDesTabella) & "}"
    End If
    FrmTab.DesTabella = MXNU.CaricaCaptionInLingua(vntDesTabella)
    FrmTab.MlngHlpTabella = lngHelp
    If IsMissing(ChiaveAgg) Then
        strChiaveAgg = ""
    Else
        strChiaveAgg = ChiaveAgg
    End If
    If IsMissing(WHEAgg) Then
        StrWheAgg = ""
    Else
        StrWheAgg = WHEAgg
    End If
    
    FrmTab.ChiaveAgg = strChiaveAgg
    FrmTab.StrWheAgg = StrWheAgg
    
    lngHwndForm = MXCT.TabellaCaricata(strNomeDbTabella)
    If lngHwndForm > 0 Then
        For ints = 0 To Forms.Count - 1
            If Forms(ints).hwnd = lngHwndForm Then
                Forms(ints).WindowState = vbNormal
                On Local Error Resume Next
                Forms(ints).SetFocus
                On Local Error GoTo 0
                Exit For
            End If
        Next
        Exit Sub
    End If

    FrmTab.Show


End Sub

Sub RiallineaProgressivi()
    Dim q%, Sql$, tB$, num As Variant
    Dim hDY As MXKit.CRecordSet
    Dim hdtb As MXKit.CRecordSet
    Dim strMsg As String
    Dim strDes As String

    metodo.MousePointer = vbHourglass
    On Local Error GoTo RPT_Err

    Sql = "SELECT * FROM TabProgressivi"
    Set hDY = MXDB.dbCreaSS(hndDBArchivi, Sql, TIPO_TABELLA)
    q = Not MXDB.dbFineTab(hDY, TIPO_DYNASET)
    strMsg = MXNU.CaricaStringaRes(2262) & Chr(10)
    Do While q
        tB$ = MXDB.dbGetCampo(hDY, TIPO_DYNASET, "NomeTabella", "")
        strDes = MXNU.CaricaCaptionInLingua(MXNU.LeggiProfilo(MXNU.File_ini_comune, "DESCRIZIONETABELLE", tB, ""))
        If strDes <> "" Then
            strMsg = strMsg & strDes & ", "
        End If
        q = MXDB.dbSuccessivo(hDY, TIPO_DYNASET)
    Loop
    q = MXDB.dbChiudiSS(hDY)
    strMsg = Left(strMsg, Len(strMsg) - 2)

    If MXNU.MsgBoxEX(strMsg, vbInformation + vbYesNo, 1007) = vbYes Then
        Set hDY = MXDB.dbCreaSS(hndDBArchivi, Sql, TIPO_TABELLA)
        q = Not MXDB.dbFineTab(hDY, TIPO_DYNASET)
        MXDB.dbBeginTrans hndDBArchivi
        Do While q
            tB$ = MXDB.dbGetCampo(hDY, TIPO_DYNASET, "NomeTabella", "")
            If tB <> "" Then
                If (Left(UCase(tB), 13) = "PACKINGCFGIMB") Or (UCase(tB) = "CONTRATTI_RCFV") Then
                    Sql = ""
                Else
                    Select Case UCase(tB)
                        Case "GESTIONEPREZZIRIGHE", "GESTIONEPREZZIRIGHETRASF"
                            Sql = "SELECT MAX(IDRiga) FROM " & tB
                        Case "TESTEENASARCO"
                            Sql = "SELECT MAX(NrBozza) FROM " & tB
                        Case "IDRIFMOVMANCDC"
                            'tB = "MovimentiCDC"
                            Sql = "SELECT MAX(IdRiferimento) FROM MovimentiCDC WHERE TipoMov=0"
                        Case "M98ARTICOLIPROD"      ' inserito per l'estensione di CALCE DEL BRENTA
                            Sql = "SELECT MAX(ID) FROM " & tB
                        Case Else
                            Sql = "SELECT MAX(Progressivo) FROM " & tB
                    End Select
                End If
                If Sql <> "" Then
                    Set hdtb = MXDB.dbCreaSS(hndDBArchivi, Sql, TIPO_TABELLA)
                    num = MXDB.dbGetCampo(hdtb, TIPO_DYNASET, 0, 0)
                    q = MXDB.dbChiudiSS(hdtb)
                    Call MXDB.dbEseguiSQL(hndDBArchivi, "UPDATE TabProgressivi SET Progr=" & num & " WHERE NomeTabella=" & hndDBArchivi.FormatoSQL(tB, DB_TEXT))
                End If
            End If
            q = MXDB.dbSuccessivo(hDY, TIPO_DYNASET)
        Loop
        MXDB.dbCommitTrans hndDBArchivi
        Call MXNU.MsgBoxEX(2121, vbInformation, 1007)
    End If
fine_rpt:
    metodo.MousePointer = vbDefault
    On Local Error GoTo 0
    q = MXDB.dbChiudiSS(hdtb)
    q = MXDB.dbChiudiSS(hDY)
    Exit Sub

RPT_Err:
    MXDB.dbRollBack hndDBArchivi
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("RiallineaProgressivi", Err.Number, Err.Description))
    Resume fine_rpt

End Sub

#If ISMETODO2005 <> 1 Then
Private Sub AbDisButtonMenu(Index As Integer, bolValue As Boolean)
    Dim i As Integer

    For i = 1 To metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus.Count
        If i = Index Then
            If bolValue Then
                metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus(i).Text = "> " & metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus(i).Text
            Else
                metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus(i).Text = Mid(metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus(i).Text, 3)
            End If
        End If
    Next i

End Sub
#End If

Public Sub Anagrafica_File(enmTipoOp As setTipoOpAn, MForm As Object, MAnagrafica As MXKit.Anagrafica, strNomeFileDef As String)
    Dim p As New PropertyBag
    Dim hf As Long
    Dim contenuto As Variant

    hf = FreeFile
    Open MXNU.PercorsoPgm & "\DEFCAMPI\" & MAnagrafica.NomeFileDef & ".DAT" For Binary As hf
    If enmTipoOp = enmCompila Then
        'Call mAnagrafica.LeggiDefAnagrafica(mForm)
        Call p.WriteProperty(strNomeFileDef, MAnagrafica)
        contenuto = p.Contents
        Put hf, 1, contenuto
    Else
        Get hf, 1, contenuto
        p.Contents = contenuto
        Set MAnagrafica = Nothing
        Set MAnagrafica = p.ReadProperty(strNomeFileDef)
        Set MAnagrafica.FormAnagr = MForm
    End If
    Close hf
    Set p = Nothing

End Sub

Private Function EseguiAgenteAzioneForm(FRM As Form, ByVal strAzione As String, bolCancellaAzione As Boolean, Optional ByVal Param As Variant) As Boolean
    bolCancellaAzione = False
    EseguiAgenteAzioneForm = True
    On Local Error Resume Next
    With FRM.MWAgt1
        If Err = 0 And .CatturaEventiForm Then
            If Not IsMissing(Param) Then .Param = Param
            EseguiAgenteAzioneForm = .Esegui(FRM, strAzione)
            bolCancellaAzione = .Cancel
        Else
            Err.Clear
        End If
    End With
    On Local Error GoTo 0
End Function

'=========================================================================================================
'           INIZIO Gestione Accessi
Sub MostraMessaggioAccessi(vntStringID As Variant, Optional vntParam As Variant)
    If (MXNU.FrmMetodo.LoadingMetodo) Then
        'sto caricando metodo -> redireziono il messaggio sulla form intro
        Call frmIntro.MostraMessaggioOperazione(vntStringID, vntParam)
    Else
        'metodo già caricato -> redireziono il messaggio sulla status bar
        Call MXNU.MostraMsgInfo(vntStringID, vntParam)
    End If
    DoEvents
End Sub

Public Function FormDefinisciAccessi(frmMyDef As Form) As Long
    If Not (frmMyDef Is Nothing) Then
        'definizione accessi form
        Set frmDefAcc.frmDef = frmMyDef
        frmDefAcc.Show vbModal
    End If
End Function
'       FINE Gestione Accessi
'=========================================================================================================

Private Function SostituisciSegnaposto(ByVal strPar$) As String
    Dim i&, strStringa$, strCar$, bolTrovataGraffa As Boolean
    bolTrovataGraffa = False
    For i = 1 To Len(strPar)
        strCar = Mid(strPar, i, 1)
        Select Case strCar
            Case ";"
                If Not bolTrovataGraffa Then
                    strStringa = strStringa & vbNullChar
                Else
                    strStringa = strStringa & strCar
                End If
            Case "{"
                bolTrovataGraffa = True
                strStringa = strStringa & strCar
            Case "}"
                bolTrovataGraffa = False
                strStringa = strStringa & strCar
            Case Else
                strStringa = strStringa & strCar
        End Select
    Next i
    SostituisciSegnaposto = strStringa
End Function

'RIF.A#9621 - gestione controllo attivo considerando le estensioni
Private Function GetActiveControl(FormAttiva As Form) As Object
Dim objActive As Object
Dim strName As String

    On Local Error Resume Next
    Set objActive = FormAttiva.ActiveControl
    strName = objActive.Name
    'è una estensione o il wrapper estensioni?
    If (strName = OGGETTO_WRAPPER_ESTENSIONE) Then
        Set objActive = objActive.object.Controls(OGGETTO_ESTENSIONE).ExtActiveControl
    ElseIf (strName = OGGETTO_ESTENSIONE) Then
        Set objActive = objActive.object.ExtActiveControl
    End If
END_GetActiveControl:
    Set GetActiveControl = objActive
    Exit Function
    
End Function



'RIF.A#4568 - aggiunto parametro ditta origine
Sub CopiaTabella(holddb As MXKit.CConnessione, hnewdb As MXKit.CConnessione, Sql$, strBlob$, NOMETABELLA$, ByVal commitparziale As Boolean, strDescrizione As String, Prc As Object, testo As Object, LungLog As Long, Optional ByVal strDittaOrigine As String = "")
    Dim hDYOld As CRecordSet, hDYNew As CRecordSet, hndSS As CRecordSet, numcampi%, nrec&, Rec&, NomeTab$, nometab1$
    Dim q%, i%, Valore As Variant, nco$, ncn$, Azz$, s$
    Dim intrans%
    Dim strNomeCampo As String
    Dim bolCampiMod As Boolean
    Dim bolOldCampiMod As Boolean
    Dim strQuery As String
    
    If Not (Prc Is Nothing) Then Prc.Value = 0
    s = LTrim$(Mid$(Sql, InStr(UCase$(Sql), "FROM") + 4))
    If InStr(s, " ") Then
        nometab1 = Left$(s, InStr(s, " ") - 1)
    Else
        nometab1 = s
    End If
    If NOMETABELLA = "" Then
        NomeTab = nometab1
    Else
        NomeTab = NOMETABELLA
    End If
    If Not (testo Is Nothing) Then
        testo.Caption = MXNU.CaricaStringaRes(1808, strDescrizione)
    End If
    s = Sql
    If InStr(UCase$(s), "WHERE") Then s = Left$(s, InStr(UCase$(s), "WHERE") - 1)
    Set hDYOld = MXDB.dbCreaDY(holddb, Sql, TIPO_TABELLA)
    'Casi Particolari di copia archivi
    Select Case UCase$(nometab1)
        Case "VISTACONTATORI"
            'Il campo Progr della VistaContatori non è aggiornabile quindi vado a scrivere
            'direttamente nella tabella contatori.
            s = swapp(s, nometab1, "TabContatori")
        Case "EXTRAMAG"
            ExtraArtSel = True
        Case "EXTRADEPOSITI"
            ExtraDepSel = True
        Case "EXTRAGIACDEPOSITI"
            ExtraGiacDepSel = True
    End Select
    Call MXNU.MsgBoxEX(MXNU.CaricaStringaRes(1812, NomeTab), vbCritical, "")
    LungLog = LungLog + Len(MXNU.CaricaStringaRes(1812, NomeTab) & ":") + 1
    DoEvents
    'LungLog = MXNU.LOFLog()
    Set hDYNew = MXDB.dbCreaDY(hnewdb, s, TIPO_TABELLA)
    q = Not MXDB.dbFineTab(hDYOld, TIPO_DYNASET)
    If q Then
        Rec = 0
        nrec = MXDB.dbNumeroRecord(hDYOld, TIPO_DYNASET)
        numcampi = MXDB.dbGetNumeroColonne(hDYOld, TIPO_DYNASET)

        ReDim mapping(0 To numcampi - 1) As Integer
        ReDim Blob(0 To numcampi - 1) As Integer
        ReDim CampiBlob(0 To numcampi - 1) As String
        Dim trovatoBlob%, numBlob%, curBlob%, tmpfile$, flen&
        For i = 0 To numcampi - 1
            nco = MXDB.dbGetNomeCampo(hDYOld, TIPO_DYNASET, i)
            mapping(i%) = MXDB.dbGetNumeroCampo(hDYNew, TIPO_DYNASET, nco)
            Blob(i%) = False
        Next i

        If Trim$(strBlob$) <> "" Then
            numBlob% = slice(strBlob$, "|", CampiBlob$())
            trovatoBlob% = (numBlob% <> 0)
            For curBlob% = 0 To numBlob% - 1
                i% = MXDB.dbGetNumeroCampo(hDYOld, TIPO_DYNASET, CampiBlob$(curBlob%))
                Blob(i%) = True
            Next curBlob%
        End If

        On Local Error GoTo err_CopiaTabella
        intrans% = False
        bolCampiMod = False
        If Not commitparziale Then MXDB.dbBeginTrans hnewdb
        While q
            If commitparziale Then MXDB.dbBeginTrans hnewdb
            intrans% = True
            q = MXDB.dbInserisci(hDYNew, TIPO_DYNASET)
            For i = 0 To numcampi - 1
                If Not Blob(i%) Then
                    strNomeCampo = MXDB.dbGetNomeCampo(hDYOld, TIPO_DYNASET, i)
                    If StrComp(strNomeCampo, "UtenteModifica", vbTextCompare) <> 0 And StrComp(strNomeCampo, "DataModifica", vbTextCompare) <> 0 Then
                        Valore = MXDB.dbGetCampo(hDYOld, TIPO_DYNASET, i, Null)
                        q = MXDB.dbSetCampo(hDYNew, TIPO_DYNASET, mapping(i), Valore)
                    Else
                        bolCampiMod = True
                    End If
                End If
            Next i
            If Not bolCampiMod Then
                bolOldCampiMod = hndDBArchivi.SettaCampiModifica
                hndDBArchivi.SettaCampiModifica = False
            End If
            If trovatoBlob% Then
                q = MXDB.dbRegistra(hDYNew, TIPO_DYNASET, "1=1")
            Else
                q = MXDB.dbRegistra(hDYNew, NO_REPOSITION, "1=1")
            End If
            If Not bolCampiMod Then
                hndDBArchivi.SettaCampiModifica = bolOldCampiMod
            End If
            
            'Registrazione dei campi BLOB (Rif.738)
            If trovatoBlob% Then
                tmpfile$ = MXNU.GetTempFile()
                For i% = 0 To numcampi - 1
                    If Blob(i%) Then
                        q% = MXDB.dbGetBlobIntoFile(hDYOld, TIPO_DYNASET, i%, tmpfile$, flen&)
                        q% = MXDB.dbSetBlobFromFile(hDYNew, TIPO_DYNASET, mapping(i%), tmpfile$, flen&)
                    End If
                Next i%
                Kill tmpfile$
            End If
            If commitparziale Then MXDB.dbCommitTrans hnewdb
            intrans% = False

            q = MXDB.dbSuccessivo(hDYOld, TIPO_DYNASET)
            Rec = Rec + 1
            i = Fix(Rec / nrec * 100)
            If Not (Prc Is Nothing) Then
                If i > Prc.Value Then Prc.Value = i
            End If
            DoEvents
        Wend

        If UCase$(nometab1) = "TIPOREGISTROIVA" Then
            'Copiando la tabella Registri Fiscali è necessario copiare anche la tabella dei
            'progressivi di stampa in quanto deve sempre corrispondere (con il codice) alla tabella Registri Fiscali.
            Dim hSS As MXKit.CRecordSet
            Dim intq As Integer
            Dim vntCodice As Variant
            Dim vntData As Variant
            
            Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT Codice,DataIniCont FROM TabEsercizi")
            intq = Not MXDB.dbFineTab(hSS)
            While intq
                vntCodice = MXDB.dbGetCampo(hSS, NO_REPOSITION, "Codice", 0)
                vntData = MXDB.dbGetCampo(hSS, NO_REPOSITION, "DataIniCont", "")
                MXDB.dbEseguiSQL hndDBArchivi, "INSERT INTO ProgressiviStampa (Esercizio,NrRegistro,DataFin,UtenteModifica,DataModifica) SELECT " & vntCodice & ",Codice,{d'" & Format$(vntData, "yyyy-mm-dd") & "'},UtenteModifica,DataModifica FROM TipoRegistroIva WHERE Codice NOT IN (SELECT NrRegistro FROM ProgressiviStampa WHERE Esercizio= " & vntCodice & ")"
                
                intq = MXDB.dbSuccessivo(hSS)
            Wend
            intq = MXDB.dbChiudiSS(hSS)
        ElseIf UCase$(nometab1) = "TABPROVVIGIONI" Then '(Rif.860)
            Set hndSS = MXDB.dbCreaSS(holddb, "SELECT CFGProvvigioni FROM ParametriGPrezzi WHERE NrRecord=1", TIPO_TABELLA)
            Valore = MXDB.dbGetCampo(hndSS, TIPO_DYNASET, "CFGProvvigioni", 0)
            q% = MXDB.dbChiudiSS(hndSS)
            MXDB.dbEseguiSQL hndDBArchivi, "UPDATE ParametriGPrezzi SET CFGProvvigioni = " & Valore & " WHERE NrRecord=1"
        ElseIf UCase$(nometab1) = "GESTIONEPREZZI" Then '(Rif.860)
            Set hndSS = MXDB.dbCreaSS(holddb, "SELECT CFGPrezzi FROM ParametriGPrezzi WHERE NrRecord=1", TIPO_TABELLA)
            Valore = MXDB.dbGetCampo(hndSS, TIPO_DYNASET, "CFGPrezzi", 0)
            q% = MXDB.dbChiudiSS(hndSS)
            MXDB.dbEseguiSQL hndDBArchivi, "UPDATE ParametriGPrezzi SET CFGPrezzi = " & Valore & " WHERE NrRecord=1"
        ElseIf (UCase$(nometab1) = "CALPRODUZIONE") Then 'RIF.A#4568 - aggiorno codice ditta sul calendario
            strQuery = "update CALPRODUZIONE" _
                & " set CODDITTA=" & hndDBArchivi.FormatoSQL(MXNU.DittaAttiva, DB_TEXT) _
                & " where CODDITTA=" & hndDBArchivi.FormatoSQL(strDittaOrigine, DB_TEXT)
            Call MXDB.dbEseguiSQL(hndDBArchivi, strQuery)
        ElseIf (UCase$(nometab1) = "PROFILOORARIOSTANDARD") Then 'RIF.A#4568 - aggiorno codice ditta sul profilo orario standard
            strQuery = "update PROFILOORARIOSTANDARD" _
                & " set CODICE=" & hndDBArchivi.FormatoSQL(MXNU.DittaAttiva, DB_TEXT) _
                & " where TIPOPO=" & MXBusiness.setTipoCalendario.tcalAziendale _
                & " and CODICE=" & hndDBArchivi.FormatoSQL(strDittaOrigine, DB_TEXT)
            Call MXDB.dbEseguiSQL(hndDBArchivi, strQuery)
        End If
        If Not commitparziale Then MXDB.dbCommitTrans hnewdb
        On Local Error GoTo 0
    End If
fine_CopiaTabella:
    On Local Error GoTo 0
    q = MXDB.dbChiudiDY(hDYNew)
    q = MXDB.dbChiudiDY(hDYOld)
Exit Sub

err_CopiaTabella:
    Dim errDescription As String
    Select Case Err.Number
       Case -2147217900  ' record già esistente
            Resume Next
       Case Else
            errDescription = Err.Description
            Call MXNU.MsgBoxEX(MXNU.CaricaStringaRes(1814, NomeTab) & "[" & errDescription & "]", vbCritical, 1007)
            If commitparziale Then
                If intrans% Then MXDB.dbRollBack hnewdb
                Resume Next
            Else
                MXDB.dbRollBack hnewdb
                Resume fine_CopiaTabella
            End If
    End Select
    
End Sub

Function EsisteCampo(strNomeTab As String, strNomeCampo As String) As Boolean
    Dim hCol As MXKit.CRecordSet
    Dim intq As Integer

    Set hCol = MXDB.dbColonne(hndDBArchivi, strNomeTab)
    Call hCol.RecSet.Find("COLUMN_NAME='" & strNomeCampo & "'", , adSearchForward, 1)
    If hCol.RecSet.EOF Or hCol.RecSet.BOF Then
        EsisteCampo = False
    Else
        EsisteCampo = True
    End If
    intq = MXDB.dbChiudiSS(hCol)

End Function


Function dbSetTextBox1(nfile As MXKit.CRecordSet, Tipo As MXKit.setTipoQuery, txtb As TextBox, prop As Proprieta_Aggiuntive) As Integer
    'eventuali restrizioni al tipo di campo vanno
    'impostate dopo la chiamata a questa funzione.
    On Local Error Resume Next
    Dim TipoCampo As setTipiDatiDB, dimcampo&
    'per compatibilità
    If prop.DataF = "" Then
        prop.DataF = txtb.DataField
    End If
    If prop.DataF = "" Then dbSetTextBox1 = False: Exit Function

    TipoCampo = MXDB.dbGetTipoCampo(nfile, Tipo, prop.DataF)
    dimcampo = MXDB.dbGetLenCampo(nfile, Tipo, prop.DataF)
    If Err Then
       dbSetTextBox1 = False
       MsgBox "Errore " & Err & ": " & Error$ & " nella funzione 'DBSetTextBox'", vbExclamation, "Errore"
    Else
        Select Case TipoCampo
           Case DB_TEXT, DB_LONGVARCHAR, DB_LONGBINARY
               txtb.MaxLength = dimcampo
               prop.Tipo = CKEY_CARASCII
               prop.frmt = "": prop.dflt = ""
           Case DB_CURRENCY
               txtb.MaxLength = 16
               'txtb.MultiLine = True
               'txtb.Alignment = 1
               prop.Tipo = CKEY_NUMFLOAT
               prop.frmt = MXNU.FORMATO_QUANTITA
               prop.dflt = 0
           Case DB_DOUBLE, DB_QUANTITA, DB_DECIMAL
               txtb.MaxLength = 16
               'txtb.MultiLine = True
               'txtb.Alignment = 1
               prop.Tipo = CKEY_NUMFLOAT
               prop.frmt = MXNU.FORMATO_QUANTITA
               prop.dflt = 0
           Case DB_SINGLE
               prop.Tipo = CKEY_NUMFLOAT
               prop.dflt = 0
           Case DB_BYTE
               txtb.MaxLength = 3
               ' txtb.MultiLine = True
               ' txtb.Alignment = 1
               prop.Tipo = CKEY_NUMINT
               prop.dflt = 0
           Case DB_INTEGER
               txtb.MaxLength = 4
               'txtb.MultiLine = True
               'txtb.Alignment = 1
               prop.Tipo = CKEY_NUMINT
               prop.dflt = 0
           Case DB_LONG
               txtb.MaxLength = 8
               'txtb.MultiLine = True
               'txtb.Alignment = 1
               prop.Tipo = CKEY_NUMINT
               prop.dflt = 0
       End Select
       dbSetTextBox1 = True
    End If
End Function


Function SpostaRec(strNomeTabella As String, strNomeCampoCod As String, ByVal Pos As Integer, varCodiceAttuale As Variant, varNuovoCodice As Variant, strWhe As String) As Boolean
    Dim intOP As Integer
    
    SpostaRec = False
    varNuovoCodice = varCodiceAttuale
    Select Case Pos
         Case BTN_PRIMO
             intOP = MXKit.dbFIND_FIRST
         Case BTN_PREC
             intOP = MXKit.dbFIND_PREVIOUS
         Case BTN_SUCC
             intOP = MXKit.dbFIND_NEXT
         Case BTN_ULTIMO
             intOP = MXKit.dbFIND_LAST
    End Select
    If MXDB.dbFind(hndDBArchivi, strNomeTabella, strWhe, strNomeCampoCod, varNuovoCodice, intOP) Then
       SpostaRec = True
    End If
    
End Function

Function Formatora(ByVal ora$, ByVal frmt$) As String

   Dim q%
   ReDim tm(5) As String
   If InStr(ora, MXNU.Sep_time) Then
      q = slice(ora, MXNU.Sep_time, tm())
   Else
      q = slice(ora, ":", tm())
   End If
   If tm(0) = "" Then tm(0) = "00"
   If tm(1) = "" Then tm(1) = "00"
   If tm(2) = "" Then tm(2) = "00"
   If frmt = MXNU.Formato_HHMMSS Then
      Formatora = Format$(tm(0), "00") & MXNU.Sep_time & Format$(tm(1), "00") & MXNU.Sep_time & Format$(tm(2), "00")
   Else
      Formatora = Format$(tm(0), "00") & MXNU.Sep_time & Format$(tm(1), "00")
   End If

End Function

