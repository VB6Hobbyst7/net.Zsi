Attribute VB_Name = "MGestMenu"
Option Explicit


Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16


Public Const ID_TLBITEM_SPACER1 = 201
Public Const ID_TLBITEM_LBLMODULO = 202
Public Const ID_TLBITEM_SPACER2 = 203
Public Const ID_TLBITEM_INFO = 204
Public Const ID_TLBITEM_PREFERENCES = 205
Public Const ID_TLBITEM_CHANGETHEME = 206

Public Const ID_TLBITEM_THEME_OFFICE2007 = 207
Public Const ID_TLBITEM_THEME_WINXPLUNA = 208
Public Const ID_TLBITEM_THEME_WINXPROYALE = 209
Public Const ID_TLBITEM_THEME_VISTA = 210
Public Const ID_TLBITEM_FRMORIGINALSIZE = 211
Public Const ID_TLBITEM_THEME_SYSTEM = 212
Public Const ID_TLBITEM_THEME_OFFICE2010 = 213

Public Const ID_TLB_PRINC = 101
Public Const ID_TLB_QUALITY = 102 '*** ATTENZIONE!! Se modificate la costante deve essere modificata anche in cAmbQualit.cls di M98Quality.dll ***
Public Const ID_TLB_DESIGNER = 103
Public Const ID_TLB_BROWSER = 104

Public Const ID_PANE_TASKBAR1 = 1001
Public Const ID_PANE_TASKBAR2 = 1002
Public Const ID_PANE_NAVBAR = 1003
Public Const ID_PANE_BROWSER = 1004
Public Const BOTTOM_PANELS_OFFSET = ID_PANE_BROWSER

Public Const ID_BAR_PROGMODULES = 1
Public Const ID_BAR_CALENDAR = 2
Public Const ID_BAR_MESSAGES = 3
Public Const ID_BAR_SHORTCUTS = 4
Public Const ID_BAR_DRAGANDRELATE As Integer = 5
Public NAVBAR_OFFSET As Integer    'ID_BAR_DRAGANDRELATE

'SottoBottoni Barra Principale
'Zoom
Public Const idxBottoneZoom100 = 50
Public Const idxBottoneZoom150 = 51
Public Const idxBottoneZoom200 = 52
'Agenti
Public Const idxBottoneNomiCtrlCmp = 54
Public Const idxBottoneSituazAnagr = 55
Public Const idxBottoneAttivaFileLog = 56
Public Const idxBottoneRicProfili = 57
'Designer
Public Const idxBottoneAttivaDesigner = 58
Public Const idxBottonePrioritaVisibVers = 59
Public Const idxBottoneCancVersione = 60
Public Const idxBottoneGrigliaDesigner = 61

Public McolMenu As CcolMenu

Public Enum setIdxToolBQuality
    idxQualityNew = 601
    idxQualityOpen = 602
    idxQualitySave = 603
    idxQualityOldDelete = 604
    idxQualityPreview = 605
    idxQualityPrint = 606
    idxQualityQuery = 607
    idxQualityList = 608
    idxQualityDeleteRow = 609
    idxQualityInsertRow = 610
    idxQualityAppendRow = 611
    idxQualitySign = 612
    idxQualitySortUp = 613
    idxQualitySortDown = 614
    idxQualityExit = 615
    idxQualityFilter = 616
    idxQualityFilterZot = 617
    idxQualityFilterWrite = 618
    idxQualityFilterClear = 619
    idxQualityDelete = 620
    idxQualityExecute = 621
    idxQualityPrintList = 622
End Enum
    
Public Enum setIdxToolBDesigner
    idxDesignerCmbDesignType = 701
    idxDesignerUnVisible = 702
    idxDesignerDisable = 703
    idxDesignerSave = 704
    idxDesignerDiscard = 705
    idxDesignerShowDiff = 706
    idxDesignerUndoCtrl = 707
    idxDesignerSetting = 708
    idxDesignerLabelSpacer1 = 709
    idxDesignerLabelTop = 710
    idxDesignerLabelLeft = 711
    idxDesignerLabelHeight = 712
    idxDesignerLabelWidth = 713
    idxDesignerLabelNoVisible = 714
    idxDesignerLabelNoEditable = 715
    idxDesignerLabelPosAltered = 716
    idxDesignerLabelTabAltered = 717
    idxDesignerSaveDesign = 730
    idxDesignerSaveVersion = 731
End Enum
    
Public McolFormsInNavBar As Collection
Public frmLog As New CGestLog

Global GBolRipristinaLayout As Boolean

Dim twMenu As MSComctlLib.TreeView
Dim ImgLMenu As MSComctlLib.ImageList

Dim colSection As Collection
Dim colValues As Collection
Dim MstrModuloCorrente As String
Dim MlngMenuID As Long
Dim MhMenuStatico As Long
Dim MstrLinguaAttuale As String
Dim MstrModuloAttuale As String   'Vedi CaricaMenuMod

Dim MSepIdx As Long
Dim MOffSetIconID As Long
Dim MBolNOAddCollMenu As Boolean

Dim mColNodesToRemove As Collection
Private Sub AddSections(f As Integer, bolPers As Boolean)
    Dim strKey As String
    Dim strValues As String
    Dim strLine As String
    strKey = ""
    strValues = ""
    Do Until EOF(f)
        Line Input #f, strLine
        '****************************************
        'Anomalia 7945
        strLine = Trim(strLine)
        strLine = Replace(strLine, vbTab, "")
        '****************************************
        If Len(Trim(strLine)) > 0 Then
            Select Case Mid(strLine, 1, 1)
                Case "["
                    If Len(strKey) = 0 Then
                        strKey = Mid(strLine, 2, InStr(2, strLine, "]") - 2)
                        If bolPers Then
                            'Tolgo il suffisso "PERS" dal nome sezione
                            If Right(UCase(strKey), 4) = "PERS" Then
                                strKey = Left(strKey, Len(strKey) - 4)
                            End If
                        End If
                    Else
                        If Not EsisteElementoCollection(colValues, "K" & strKey) Then
                            Call colValues.Add(strValues, "K" & strKey)
                        Else
                            If bolPers Then
                                strValues = colValues("K" & strKey) & strValues
                                colValues.Remove "K" & strKey
                                Call colValues.Add(strValues, "K" & strKey)
                            End If
                        End If
                        If Not EsisteElementoCollection(colSection, "K" & strKey) Then
                            Call colSection.Add(strKey, "K" & strKey)
                        End If
                        strKey = Mid(strLine, 2, InStr(2, strLine, "]") - 2)
                        If bolPers Then
                            'Tolgo il suffisso "PERS" dal nome sezione
                            If Right(UCase(strKey), 4) = "PERS" Then
                                strKey = Left(strKey, Len(strKey) - 4)
                            End If
                        End If
                        strValues = ""
                    End If
                Case ";"
                
                Case Else
                    If Len(strKey) > 0 Then
                        strValues = strValues & strLine & vbCrLf
                    End If
            End Select
        End If
    Loop

End Sub

'Cambia la dll delle risorse in lingua per gli oggetti CodeJock
Public Sub CambiaRisorseCJ()
    CommandBarsGlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"
    Call metodo.CreateToolBars  'Affinchè le risorse in lingua sulla toolbar codejock si aggiornino, è necessario ricreare tutti i pulsanti (!?!)
    DockingPaneGlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"
    ShortcutBarGlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"
    frmModuli2005.CambiaRisorse
    frmModuli.CambiaRisorse
    metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Title = frmModuli2005.ShortcutBar1.Selected.Caption
    If Not (mMessagingEngine Is Nothing) Then
        Call mMessagingEngine.ChangeCJResources
        metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR2).Title = MXNU.CaricaStringaRes(11991)
    End If
    metodo.CommandBars.RecalcLayout
    Select Case LCase(GTemaAttivo)
        Case "office2007"
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2007).DefaultItem = True
        Case "winxp.luna"
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_WINXPLUNA).DefaultItem = True
        Case "winxp.royale"
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_WINXPROYALE).DefaultItem = True
        Case "vista"
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_VISTA).DefaultItem = True
        Case "system"
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_SYSTEM).DefaultItem = True
        Case Else
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2007).DefaultItem = True
    End Select
End Sub

'Private Sub CaricaMenuComuniDaColl()
'    Dim cm As CGestMenu
'    Dim Control As CommandBarControl
'    Dim Bar As CommandBarPopup
'    Dim BarParent As CommandBarPopup
'    Dim Btn As CommandBarControl
'    Dim ControlWindow As CommandBarPopup, ControlHelp As CommandBarPopup
'
'    For Each cm In McolMenu
'        If cm.Modulo = "Comuni" Then
'            If cm.key = "Menu" Then
'                Set BarParent = metodo.CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, cm.Caption, 1, False)
'            ElseIf cm.key = "Aiuto" Then
'                Set BarParent = metodo.CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, cm.Caption, -1, False)
'            ElseIf cm.key <> "Strum" Or cm.ParentKey <> "Strum" Then
'                Set Btn = BarParent.CommandBar.Controls.Add(xtpControlPopup, cm.MenuID, cm.Caption, , False)
'            End If
'        End If
'    Next
'
'End Sub

Private Function CaricataFormPrelievo() As Boolean
    Dim frm As Form
    
    CaricataFormPrelievo = False
    On Local Error Resume Next
    For Each frm In VB.Forms
        If frm.Name = "frmPrelDoc" Then
            CaricataFormPrelievo = True
            Exit For
        End If
    Next

End Function

Private Sub GestMnuMenu(ByVal indice As Integer)
     Select Case indice
        Case 0
            'Anomalia #6844 + Anom. Evolus #91 - Inibisce il cambio ditta se è aperta un'anteprima di stampa
            If MXCREP.StampaInCorso Then
                Call MXNU.MsgBoxEX(1096, vbOKOnly + vbExclamation, 1007)
                Exit Sub
            End If
            'Anomalia nr. 9790
            If CaricataFormPrelievo Then Exit Sub
            InSelezioneDitta = True
            Call ApriDittaAnno(False, "", "")
            InSelezioneDitta = False
        '             If ApriDittaAnno(False, "", "") Then
        '                Call MDIForm_Resize
        '             End If
        Case 1
            'Anomalia #6844 + Anom. Evolus #91 - Inibisce il cambio esercizio se è aperta un'anteprima di stampa
            If MXCREP.StampaInCorso Then
                Call MXNU.MsgBoxEX(1096, vbOKOnly + vbExclamation, 1007)
                Exit Sub
            End If
            'Anomalia nr. 9790
            If CaricataFormPrelievo Then Exit Sub
            Dim NuovoAnno As Integer
            If SelezioneAnno(True, NuovoAnno) Then
                Call ChiudiFormAttive
                MXNU.AnnoAttivo = NuovoAnno   'Rif. Sk. Anomalie Nr. 5113
                Call ApriAnno(True)
            End If
        Case 3
           metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneModuli).Execute
        '            #If ISMETODOXP = 1 Then
        '                If MXNU.MetodoXP Then
        '                    If frmModuli.StatoMenuAlbero(0) = False Then
        '                        'il menu moduli non e' aperto ne lockato
        '                        frmModuli.StatoMenuModuli = MnuModATTIVO
        '                        Call frmModuli.AlberoBloccatoMenu(True, 0)
        '                    Else
        '                        'il menu moduli e' gia' aperto e lockato
        '                        'frmModuli.StatoMenuModuli = MnuModNONATTIVO
        '                        Call frmModuli.AlberoBloccatoMenu(False, 0)
        '                    End If
        '                Else
        '                    SetCursorPos 10, ((metodo.Height - metodo.ScaleHeight) + 30) / Screen.TwipsPerPixelY
        '                    Call AttivaMenuMetodo
        '                End If
        '
        '            #Else
        '                 SetCursorPos 10, ((metodo.Height - metodo.ScaleHeight) + 30) / Screen.TwipsPerPixelY
        '                 Call AttivaMenuMetodo
        '
        '            #End If
        
        Case 5
            'Anomalia #6844 + Anom. Evolus #91 - Inibisce il cambio ditta se è aperta un'anteprima di stampa
            If MXCREP.StampaInCorso Then
                Call MXNU.MsgBoxEX(1096, vbOKOnly + vbExclamation, 1007)
                Exit Sub
            End If
            
            'Rif. anomalia XP #3160
            If MXNU.LoginIntegrato Then
                Call MXNU.MsgBoxEX(3015, vbExclamation, 1007)
            Else
                Call CambioUtenteAttivo
            End If
            'Call MDIForm_Resize
        Case 6
        Case 8 'Uscita
            Unload metodo
    End Select

End Sub

Public Sub LoadIniMenu(twMenuModuli As MSComctlLib.TreeView)

Dim strFile As String
Dim strFileMenu As String
Dim strFileMenuPers As String
Dim strLine As String
Dim strValue As String, strKey As String, strValues As String
Dim strSection As String
Dim strItem As Variant
Dim objNode As Node
'Dim t As Double
Dim f As Integer
Dim bolAbilita As Boolean
Dim i As Long
Dim strTempDir As String
Dim objMenu As CGestMenu

'    t = Timer
    Call InizializzaLetturaAccessi(MXNU.UtenteAttivo, True)
    Set twMenu = twMenuModuli
    Set ImgLMenu = twMenuModuli.ImageList
    metodo.CommandBars.AddImageList twMenuModuli.ImageList
    
    If Not (colSection Is Nothing) Then
        While colSection.Count() > 0
            Call colSection.Remove(1)
        Wend
    Else
        Set colSection = New Collection
    End If
    If Not (colValues Is Nothing) Then
        While colValues.Count() > 0
            Call colValues.Remove(1)
        Wend
    Else
        Set colValues = New Collection
    End If
    
    If Not (McolMenu Is Nothing) Then
        Set McolMenu = Nothing
    End If
    Set McolMenu = New CcolMenu
    
    Call MostraMessaggioAccessi(9013)
    
    strFileMenu = MXNU.NomeMenu
    strFileMenuPers = "menu" & MXNU.LinguaAttiva & ".ini"
    MXNU.File_ini_Menu = MXNU.PercorsoPgm & "\" & strFileMenu
    strTempDir = MXNU.GetTempDir
    
    On Local Error Resume Next
    Kill strTempDir & NOME_FILE_MENU_TMP
    Kill strTempDir & NOME_FILE_MENUPERS_TMP
    Kill strTempDir & NOME_FILE_MENUPERSDITTA_TMP
    
    FileCopy MXNU.File_ini_Menu, strTempDir & NOME_FILE_MENU_TMP
    FileCopy MXNU.PercorsoPers & "\" & strFileMenuPers, strTempDir & "\" & NOME_FILE_MENUPERS_TMP
    FileCopy MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\" & strFileMenuPers, strTempDir & "\" & NOME_FILE_MENUPERSDITTA_TMP
    
    strFileMenu = MXNU.File_ini_Menu
    MXNU.File_ini_Menu = strTempDir & NOME_FILE_MENU_TMP
    f = FreeFile
    Open MXNU.File_ini_Menu For Input Access Read Shared As #f
    Call AddSections(f, False)
    Close #f
    
    'Menu Pers
    strFileMenu = strFileMenuPers
    'Anomalia 7811 e 8688 - Considero entrambi i file di pers e persditta
    strFileMenu = CercaDirFile(strFileMenu, MXNU.PercorsoPers)
    If strFileMenu <> "" Then
        f = FreeFile
        Open strFileMenu For Input Access Read Shared As #f
        Call AddSections(f, True)
        Close #f
    End If
    
    strFileMenu = strFileMenuPers
    strFileMenu = CercaDirFile(strFileMenu, MXNU.PercorsoPers & "\" & MXNU.DittaAttiva)
    If strFileMenu <> "" Then
        f = FreeFile
        Open strFileMenu For Input Access Read Shared As #f
        Call AddSections(f, True)
        Close #f
    End If
    
    twMenu.Nodes.Clear
    twMenu.Visible = False
    Set objNode = twMenu.Nodes.Add(, tvwFirst, "Metodo98", "MetodoEvolus", "Metodo98")
    objNode.Expanded = True
    
    Set objMenu = New CGestMenu
    With objMenu
        .Caption = "MetodoEvolus"
        .key = "Metodo98"
        .HasSeparator = False
        .HelpContextID = 0
        .ParentKey = ""
        .Modulo = "Metodo98"
        .MenuID = 0
    End With
    McolMenu.Add objMenu
    
    MlngMenuID = 1000   'Gli ID devono essere univoci tra Menu e Toolbar varie, quindi parto da un valore superiore
    Call LoadNode("MODULI", objNode, True)
    DoEvents
    Call LoadNode("MODULIPERS", objNode, True)
    
    twMenu.Visible = True
'    MsgBox "Lettura INI: " & Timer - t
    Call MostraMessaggioAccessi(9014)
'    t = Timer
    Call InizializzaBufferAccessi
    
    Set mColNodesToRemove = New Collection
    Call ImpostaAccessi(twMenu.Nodes(1))
    'rimovo i nodi
    On Local Error Resume Next
    For i = 1 To mColNodesToRemove.Count
        twMenu.Nodes.Remove mColNodesToRemove(i)
    Next
    
'    Dim lngNNodi As Long, vntIndex As Variant, vntNome As String
'    lngNNodi = twMenu.Nodes.Count
'    For i = 1 To lngNNodi
'        On Local Error Resume Next
'        'Aggiunto controllo con IIF (Anomalia 7608)
'        If twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key = McolMenu(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key).Modulo Then
'            bolAbilita = (ChiaveDammiAccessi(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key, ""))
'            If bolAbilita And MXNU.CtrlAccessi Then
'                vntNome = LCase(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key)
'                'If LCase(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key) <> "metodo98" Then
'                Select Case LCase(vntNome)
'                    'Se sto analizzando il modulo root dell'albero ("Metodo98") o il modulo comuni lo abilito sempre, altrimenti non si vedranno i moduli abilitati per l'utente
'                    Case "metodo98", "comuni", "menu", "aiuto"
'                        bolAbilita = True
'                    Case Else
'                        bolAbilita = LeggiAccessi(MXNU.UtenteAttivo, McolMenu(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key).HelpContextID, ID_SCHEDA_FORM)
'                End Select
'            End If
'            If bolAbilita Then
'                bolAbilita = MXNU.AssegnaModuliChiave(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key, "")
'                If Not bolAbilita Then
'                    twMenu.Nodes.Remove twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key
'                End If
'            Else
'                twMenu.Nodes.Remove twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key
'            End If
'        Else
'            If LCase(Left(McolMenu(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key).key, 4)) = "menu" Or _
'               LCase(Left(McolMenu(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key).key, 5)) = "aiuto" Then
'                bolAbilita = True
'            Else
'                bolAbilita = MenuDefinisciAccessi(McolMenu(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key))
'            End If
'            If bolAbilita Then
'                '***TSE controllo moduli INTERNI (Analitica,finanziaria,ecc..)
'                On Local Error Resume Next
'                vntIndex = ""
'                vntNome = "mnu" & ScomponiChiave(McolMenu(twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key).key, vntIndex)
'                If (Err <> 0) Then vntIndex = ""
'                On Local Error GoTo 0
'                bolAbilita = MXNU.AssegnaModuliChiave(vntNome, vntIndex)
'                If Not bolAbilita Then
'                    twMenu.Nodes.Remove twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key
'                End If
'            Else
'                twMenu.Nodes.Remove twMenu.Nodes(IIf(i - (lngNNodi - twMenu.Nodes.Count) > 0, i - (lngNNodi - twMenu.Nodes.Count), (lngNNodi - twMenu.Nodes.Count) - i)).key
'            End If
'        End If
'    Next i
    
    Call TerminaLetturaAccessi
    Set colValues = Nothing
    Set colSection = Nothing
'    MsgBox "Lettura Chiave+Accessi: " & Timer - t
    
#If ISNUCLEO = 0 Then
    'rif.S#1358 - caricamento menu MetodoOLAP
    Call MOLAP_LoadMenu
#End If
    
    
    'nascondo menu "Menù" sull'albero  (vedi vecchia CaricaModuli su frmModuli)  '<<< Non possibile su Evolus: non c'è un menu statico sulla MDI a differenza di MetodoXP. Il menu Menu deve esserci anche nell'albero
    'On Local Error Resume Next
    'twMenu.Nodes.Remove "Menu"
    'On Local Error GoTo 0

    
    Call MostraMessaggioAccessi("")
    Set mColNodesToRemove = Nothing
End Sub



#If ISNUCLEO = 0 Then
'----------------------------------------------- rif.S#1358 - caricamento menu MetodoOLAP -----------------------------------------------
'------------------------------------------------------------
'nome:          MOLAP_LoadMenu
'descrizione:   caricamento dei menu MetodoOLAP
'parametri:
'annotazioni:
'------------------------------------------------------------
Private Function MOLAP_LoadMenu() As Boolean
Dim bolRes As Boolean
Dim rootMenu As CGestMenu
Dim rootNode As MSComctlLib.Node
Dim objBrioFactory As Object
Dim bolHardwareKeyPresent As Boolean

    bolRes = True
    Call MostraMessaggioAccessi(9009, Array("", "Metodo Olap"))
    'caricamento brio factory
    On Local Error Resume Next
    Set objBrioFactory = CreateObject("MxBrioFactory.cBrioFactory")
    If (objBrioFactory Is Nothing) Then
        bolRes = False
        GoTo MOLAP_LoadMenu_END
    Else
        'controllo presenza della chiave MetodoOLAP
        bolHardwareKeyPresent = (StrComp(objBrioFactory.GetMolapData("HARDWAREKEYPRESENT"), "notpresent", vbTextCompare) <> 0)
        If (bolHardwareKeyPresent) Then
            bolHardwareKeyPresent = (StrComp(objBrioFactory.GetMolapData("TERMINALSERVERSTATUS"), "tsecheckpassed", vbTextCompare) = 0)
        End If
        'caricamento del menu
        If (bolHardwareKeyPresent) Then
            On Local Error GoTo MOLAP_LoadMenu_ERR
            Set mColAzioniMetodoOLAP = New Collection
            'creazione menu principale
            Set rootMenu = New CGestMenu
            With rootMenu
                .Caption = "Metodo OLAP"
                .HasSeparator = False
                .key = "MetodoOLAP"
                .MenuID = MlngMenuID: MlngMenuID = MlngMenuID + 1
                .Modulo = "MetodoOLAP"
                .ParentKey = "Metodo98"
            End With
            Call McolMenu.Add(rootMenu)
            
            Set rootNode = twMenu.Nodes.Add("Metodo98", tvwChild, rootMenu.key, rootMenu.Caption, "MOLAP")
            rootNode.Tag = 0
            rootNode.EnsureVisible
            'caricamento del menu Configurazione
            Call MOLAP_LoadConfigurationMenu(rootMenu)
            'caricamento dei sotto-moduli
            Call MOLAP_LoadOlapCategories(rootMenu, objBrioFactory)
        End If
    End If

MOLAP_LoadMenu_END:
    Set rootMenu = Nothing
    Set rootNode = Nothing

    MOLAP_LoadMenu = bolRes
    On Local Error GoTo 0
    Exit Function

MOLAP_LoadMenu_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("MOLAP_LoadMenu", lngErrCod, strErrDsc))
    Resume MOLAP_LoadMenu_END
Resume
End Function

'------------------------------------------------------------
'nome:          MOLAP_LoadConfigurationMenu
'descrizione:   caricamento del menu configurazione
'parametri:     menu padre
'annotazioni:
'------------------------------------------------------------
Private Sub MOLAP_LoadConfigurationMenu(ByVal objParent As CGestMenu)
Dim bolRes As Boolean
Dim lngErrCod As Long
Dim strErrDsc As String
Dim objMenu As CGestMenu
Dim objNode As MSComctlLib.Node
Dim objMenuItem As CGestMenu


    bolRes = True
    On Local Error GoTo MOLAP_LoadConfigurationMenu_ERR

    Set objMenu = New CGestMenu
    With objMenu
        .Caption = "Configurazione"
        .HasSeparator = False
        .HelpContextID = 0
        .key = ComponiChiave(METODO_OLAP_MENU_KEY & "Config", 0)
        .MenuID = MlngMenuID: MlngMenuID = MlngMenuID + 1
        .Modulo = "MetodoOLAP"
        .ParentKey = objParent.key
        
        'generazione nodo albero
        Set objNode = twMenu.Nodes.Add(objParent.key, tvwChild, .key, .Caption, "CartCh")
        objNode.Tag = 0
    End With
    McolMenu.Add objMenu

    'caricamento item configurazione
    Set objMenuItem = New CGestMenu
    With objMenuItem
        .Caption = "Opzioni"
        .HasSeparator = False
        .HelpContextID = 0
        .key = ComponiChiave(METODO_OLAP_MENU_KEY & "ConfigItem", 0)
        .MenuID = MlngMenuID: MlngMenuID = MlngMenuID + 1
        .Modulo = "MetodoOLAP"
        .ParentKey = objMenu.key
        
        'generazione nodo albero
        Set objNode = twMenu.Nodes.Add(objMenu.key, tvwChild, .key, .Caption, "Exe")
        objNode.Tag = 0
    End With
    McolMenu.Add objMenuItem

MOLAP_LoadConfigurationMenu_END:
    Set objMenu = Nothing
    Set objMenuItem = Nothing
    Set objNode = Nothing
    'Chiusura oggetti, dynaset
    On Local Error GoTo 0
    If Not bolRes Then Err.Raise lngErrCod, "MOLAP_LoadConfigurationMenu", strErrDsc
    Exit Sub

MOLAP_LoadConfigurationMenu_ERR:
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Resume MOLAP_LoadConfigurationMenu_END
End Sub

'------------------------------------------------------------
'nome:          MOLAP_LoadOlapCategories
'descrizione:   caricamento dei moduli di metodo olap
'parametri:     modulo root
'annotazioni:
'------------------------------------------------------------
Private Sub MOLAP_LoadOlapCategories(ByVal objModule As CGestMenu, ByVal objBrioFactory As Object)
Dim bolRes As Boolean
Dim lngErrCod As Long
Dim strErrDsc As String
Dim objCategory As CGestMenu
Dim objNode As MSComctlLib.Node
Dim XmlLicences As MSXML2.DOMDocument
Dim bfAnswer As String
Dim oCategoryNodeList As MSXML2.IXMLDOMNodeList
Dim oCategoryNode As MSXML2.IXMLDOMNode
Dim nodeIndex As Integer
Dim categoryName As String
Dim cConfig As cMOlapConfig

    bolRes = True
    On Local Error GoTo MOLAP_LoadOlapModels_ERR
    'RIF.A#8327 - caricamento della configurazione
    Set cConfig = New cMOlapConfig
    Call cConfig.LoadCurrentConfig

    'richiesta all'oggetto brio factory del file licenze
    bfAnswer = objBrioFactory.GetMolapData("MODULES&" & cConfig.PercorsoModelli)
    If (Len(bfAnswer) > 0) Then
        Set XmlLicences = New MSXML2.DOMDocument
        If (XmlLicences.loadXML(bfAnswer)) Then
            Set oCategoryNodeList = XmlLicences.selectNodes("models/category")
            nodeIndex = -1
            For Each oCategoryNode In oCategoryNodeList
                categoryName = MXUTil.XmlGetAttributeValue(oCategoryNode, "name")
                Set objCategory = New CGestMenu
                nodeIndex = nodeIndex + 1
                With objCategory
                    .Caption = categoryName
                    .HasSeparator = False
                    .HelpContextID = 0
                    .key = ComponiChiave(METODO_OLAP_MENU_KEY & categoryName, nodeIndex)
                    .MenuID = MlngMenuID: MlngMenuID = MlngMenuID + 1
                    .Modulo = "MetodoOLAP"
                    .ParentKey = objModule.key
                    
                    'generazione nodo albero
                    Set objNode = twMenu.Nodes.Add(objModule.key, tvwChild, .key, .Caption, "OLAPCART")
                    objNode.Tag = 0
                End With
                McolMenu.Add objCategory
                'caricamento dei moduli
                Call MOLAP_LoadOlapModels(oCategoryNode, objCategory, objBrioFactory)
            Next
        End If
    End If

MOLAP_LoadOlapModels_END:
    'Chiusura oggetti, dynaset
    On Local Error GoTo 0
    If Not bolRes Then Err.Raise lngErrCod, "MOLAP_LoadOlapModels", strErrDsc
    Exit Sub

MOLAP_LoadOlapModels_ERR:
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Resume MOLAP_LoadOlapModels_END
Resume
End Sub

'------------------------------------------------------------
'nome:          MOLAP_LoadOlapModels
'descrizione:   caricamento dei moduli olap
'parametri:     nodo xml che rappresenta la categoria
'               modulo root
'               menu categoria
'annotazioni:
'------------------------------------------------------------
Private Sub MOLAP_LoadOlapModels(ByVal oCategoryNode As MSXML2.IXMLDOMNode, objCategory As CGestMenu, objBrioFactory As Object)
Dim bolRes As Boolean
Dim lngErrCod As Long
Dim strErrDsc As String
Dim oModelNodeList As MSXML2.IXMLDOMNodeList
Dim oModelNode As MSXML2.IXMLDOMNode
Dim objMenuItem As CGestMenu
Dim objNode As MSComctlLib.Node
Dim nodeIndex As Integer
Dim bolModulePresent As Boolean
Dim categoryName As String
Dim modelName As String
Dim modelDescription As String
Dim isDataModel As Boolean

    bolRes = True
    On Local Error GoTo MOLAP_LoadOlapModels_ERR
    Set oModelNodeList = oCategoryNode.selectNodes("model")
    nodeIndex = -1
    For Each oModelNode In oModelNodeList
        categoryName = MXUTil.XmlGetAttributeValue(oCategoryNode, "name", vbNullString)
        modelName = MXUTil.XmlGetAttributeValue(oModelNode, "name", vbNullString)
        modelDescription = MXUTil.XmlGetAttributeValue(oModelNode, "description", vbNullString)
        If (Len(modelDescription) = 0) Then modelDescription = modelName
                'caricamento item configurazione
                Set objMenuItem = New CGestMenu
                nodeIndex = nodeIndex + 1
                With objMenuItem
                    .Caption = modelDescription
                    .HasSeparator = False
                    .HelpContextID = 0
                    .key = ComponiChiave(objCategory.key & "_Item", nodeIndex)
                    .MenuID = MlngMenuID: MlngMenuID = MlngMenuID + 1
                    .Modulo = "MetodoOLAP"
                    .ParentKey = objCategory.key
                    'generazione nodo albero
                    Set objNode = twMenu.Nodes.Add(objCategory.key, tvwChild, .key, .Caption, "OLAPITEM")
                    objNode.Tag = 0
                End With
                McolMenu.Add objMenuItem
                
                'aggiunta percorso bqy
                Call mColAzioniMetodoOLAP.Add(categoryName & "\" & modelName & ".bqy", objMenuItem.key)
    Next

MOLAP_LoadOlapModels_END:
    'Chiusura oggetti, dynaset
    On Local Error GoTo 0
    If Not bolRes Then Err.Raise lngErrCod, "MOLAP_LoadOlapModels", strErrDsc
    Exit Sub

MOLAP_LoadOlapModels_ERR:
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Resume MOLAP_LoadOlapModels_END
Resume
End Sub

#End If


Private Function ImpostaAccessi(ByVal oTreeNode As MSComctlLib.Node) As Boolean
Dim bolAbilita As Boolean
Dim oChildNode As MSComctlLib.Node

    'abilito sempre il menu principale
    If (oTreeNode Is twMenu.Nodes(1)) Then
        bolAbilita = True
    ElseIf (oTreeNode.Parent Is twMenu.Nodes(1)) Then
        'modulo
        bolAbilita = ImpostaAccessiModulo(oTreeNode)
    Else
        'voce di menu
        Select Case LCase(oTreeNode.key)
            Case "menu", "aiuto": bolAbilita = True   'Anomalia Evolus 94
            Case Else
                'Anomalia Evolus 94
                If InStr(LCase(oTreeNode.key), "menuitem") > 0 Or InStr(LCase(oTreeNode.key), "aiutoitem") > 0 Then
                    bolAbilita = True
                Else
                    bolAbilita = ImpostaAccessiMenu(oTreeNode)
                End If
        End Select
    End If
    'ricorsione sui figli
    If (bolAbilita) Then
        Set oChildNode = oTreeNode.Child
        Do While (Not oChildNode Is Nothing)
            Call ImpostaAccessi(oChildNode)
            Set oChildNode = oChildNode.Next
        Loop
    End If
End Function


Private Function ImpostaAccessiModulo(ByVal oTreeNode As MSComctlLib.Node) As Boolean
Dim bolAbilita As Boolean

    'Anomalia 9014
    If (LCase(oTreeNode.key) <> "metodo98") And (LCase(oTreeNode.key) <> "comuni") Then
        bolAbilita = MXNU.AssegnaModuliChiave(oTreeNode.key, vbNullString)
    Else
        bolAbilita = True
    End If
    If (bolAbilita And MXNU.CtrlAccessi) Then
        If (LCase(oTreeNode.key) <> "metodo98") And (LCase(oTreeNode.key) <> "comuni") Then   'Aggiunto anche modulo Comuni (Anomalia Evolus Nr. 94)
            'Se sto analizzando il modulo root dell'albero ("Metodo98") lo abilito sempre, altrimenti non si vedranno i moduli abilitati per l'utente
            bolAbilita = LeggiAccessi(MXNU.UtenteAttivo, McolMenu(oTreeNode.key).HelpContextID, ID_SCHEDA_FORM)
        End If
    End If
    If (bolAbilita) Then
        bolAbilita = MXNU.AssegnaModuliChiave(oTreeNode.key, "")
    End If
    'se non abilitato => rimuovo il nodo e tutti i suoi figli
    'NOTA: uso una collection perchè altrimenti rimuovendo direttamente i nodi va in errore il Node.Next
    If (Not bolAbilita) Then
        mColNodesToRemove.Add oTreeNode.key
    End If

    ImpostaAccessiModulo = bolAbilita
End Function


Private Function ImpostaAccessiMenu(ByVal oTreeNode As MSComctlLib.Node) As Boolean
Dim bolAbilita As Boolean
Dim vntIndex As Variant
Dim vntNome As Variant

    bolAbilita = MenuDefinisciAccessi(McolMenu(oTreeNode.key))
    If (bolAbilita) Then
        '***TSE controllo moduli INTERNI (Analitica,finanziaria,ecc..)
        On Local Error Resume Next
        vntIndex = ""
        vntNome = "mnu" & ScomponiChiave(McolMenu(oTreeNode.key).key, vntIndex)
        If (Err <> 0) Then vntIndex = ""
        On Local Error GoTo 0
        bolAbilita = MXNU.AssegnaModuliChiave(vntNome, vntIndex)
    End If
    'se non abilitato => rimuovo il nodo e tutti i suoi figli
    'NOTA: uso una collection perchè altrimenti rimuovendo direttamente i nodi va in errore il Node.Next
    If (Not bolAbilita) Then
        mColNodesToRemove.Add oTreeNode.key
    End If
    
    ImpostaAccessiMenu = bolAbilita
End Function


Public Sub AttivaVoceMenu(ByVal NodeKey As Variant)
'    TrwModuli.SelectedItem = TrwModuli.Nodes(NodeKey)
'    TrwModuli_DblClick
    Dim Nome As Variant, indice As Variant
    On Local Error Resume Next   'Aggiunto resume next x anomalia
    Nome = ScomponiChiave(McolMenu(NodeKey).key, indice)
    If Nome <> "MenuItem" Then
        'Call EseguiAzione_I(Nome, Val(Indice), McolMenu(NodeKey).HelpContextID)
        'Anomalia 9049
        Call EseguiAzione_i(Nome, Val(indice), McolMenu(frmMenu.TrwModuli.Nodes(NodeKey).key).HelpContextID)
    Else
        Call GestMnuMenu(Val(indice))
    End If
    On Local Error GoTo 0

End Sub

Public Sub CaricaMenuXCB(ByVal NomeModulo As String)
    Dim objMenu As CGestMenu
    Dim objNode As Node
    Dim i As Long
    Dim Control As CommandBarControl
    Dim Bar As CommandBarPopup
    Dim ControlWindow As CommandBarPopup, ControlHelp As CommandBarPopup

    metodo.CommandBars.ActiveMenuBar.Controls.DeleteAll

    If LCase(NomeModulo) <> "metodo98" Then
        If twMenu.Nodes(NomeModulo).children > 0 Then
            Set objNode = twMenu.Nodes(NomeModulo).Child
            Do While Not (objNode Is Nothing)
                'Set Bar = metodo.CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, objNode.text, -1, False)
                'Uso la proprietà Caption della collection invece del testo del nodo dell'albero, altrimenti non si vedono le lettere sottolineate per i menù di primo livello
                Set Bar = metodo.CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, McolMenu(objNode.key).Caption, -1, False)
                Call CaricaSubMenuXCB(objNode.key, Bar)
                Set objNode = objNode.Next
            Loop
        End If
'    Else
'        If MXNU.CtrlAccessi Then
'            On Local Error Resume Next
'            Set objNode = twMenu.Nodes("Menu")
'            If Err.Number <> 0 Then
'                Call CaricaMenuComuniDaColl
'            End If
'            On Local Error GoTo 0
'            'MBolNOAddCollMenu = True
'            'Call LoadNode("Comuni", twMenu.Nodes(1), True)
'            'MBolNOAddCollMenu = False
'        End If
    End If

    If NomeModulo <> "Comuni" Then
        Set objNode = twMenu.Nodes("Menu")
        If Not (objNode Is Nothing) Then
            Set Bar = metodo.CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, McolMenu(objNode.key).Caption, 1, False)
            Call CaricaSubMenuXCB(objNode.key, Bar)
            metodo.CommandBars.KeyBindings.Add FCONTROL, Asc("L"), McolMenu("MenuItem_3").MenuID
        End If
        Set objNode = twMenu.Nodes("Aiuto")
        If Not (objNode Is Nothing) Then
            Set Bar = metodo.CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, McolMenu(objNode.key).Caption, -1, False)
            Call CaricaSubMenuXCB(objNode.key, Bar)
            Bar.CommandBar.Controls.Add xtpControlButton, 35000, "&Window List", -1, False
        End If
    'ElseIf NomeModulo = "Aiuto" Then
    '    mnuWndList = New MdiWindowListItem("mnuWndList", "Finestre &Aperte")
    '    dotNetBarManager1.Bars("mnuMain").Items("Aiuto").SubItems.Add (mnuWndList)
    End If
    metodo.CommandBars.RecalcLayout
    'metodo.CommandBars.DockToolBar metodo.CommandBars.ActiveMenuBar, 0, 0, xtpBarLeft

    On Local Error Resume Next
    Dim ScreenSizeX As Long, ScreenSizeY As Long
    ScreenSizeX = (Screen.Width \ Screen.TwipsPerPixelX)
    ScreenSizeY = (Screen.Height \ Screen.TwipsPerPixelY)

'    #If TOOLS <> 1 Then
'        'sfondo personalizzato per versione
'        metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo." & MXNU.VersioneAttiva & ".jpg")
'        If (Err <> 0) Then
'            Err.Clear
'            Set metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo." & MXNU.VersioneAttiva & ".bmp")
'        End If
'        'sfondo personalizzato per utente
'        If (Err <> 0) Then
'            Err.Clear
'            metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo." & MXNU.UtenteAttivo & ".jpg")
'        End If
'        If (Err <> 0) Then
'            Err.Clear
'            Set metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo." & MXNU.UtenteAttivo & ".bmp")
'        End If
'        'sfondo standard per risoluzione video
'        If (Err <> 0) Then
'            Err.Clear
'            metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo" & ScreenSizeX & "x" & ScreenSizeY & ".jpg")
'        End If
'        If (Err <> 0) Then
'            Err.Clear
'            Set metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo.jpg")
'        End If
'        If (Err <> 0) Then
'            Err.Clear
'            Set metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo.bmp")
'        End If
'    #Else
'        Err.Clear
'        'Sfondo Tools per risoluzione video
'        metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondoTools" & ScreenSizeX & "x" & ScreenSizeY & ".jpg")
'        If (Err <> 0) Then
'            Err.Clear
'            metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondoTools.jpg")
'        End If
'        If (Err <> 0) Then
'            Err.Clear
'            Set metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondoTools.bmp")
'        End If
'    #End If
    
End Sub


Public Sub ScaricaMenu()
    Set colSection = Nothing
    Set colValues = Nothing
    Set McolMenu = Nothing
End Sub



Public Sub AggiungiImagesXCB()
'    Dim i As Long
'    For i = 1 To metodo.ImglistBottoniXP.ListImages.Count
'        metodo.CommandBars.AddIconHandle metodo.ImglistBottoniXP.ListImages(i).Picture.handle, i, 0, False
'    Next i
'    MOffSetIconID = metodo.ImglistBottoniXP.ListImages.Count + 1
'    For i = 1 To frmModuliXP.ImgLstModuli.ListImages.Count
'        metodo.CommandBars.AddIconHandle frmModuliXP.ImgLstModuli.ListImages(i).Picture.handle, i + MOffSetIconID, 0, False
'    Next i
    
End Sub

Private Sub CaricaSubMenuXCB(ByVal NomeModulo As String, BarParent As CommandBarPopup)
    Dim objNode As Node
    
    Dim Btn As CommandBarControl
    Dim bolPrimo As Boolean
    Static bolAddSeparator As Boolean
    
    With BarParent.CommandBar.Controls
        If twMenu.Nodes(NomeModulo).children > 0 Then
            Set objNode = twMenu.Nodes(NomeModulo).Child
            Do While Not (objNode Is Nothing)
                If objNode.children > 0 Then
                    Set Btn = .Add(xtpControlPopup, McolMenu(objNode.key).MenuID, McolMenu(objNode.key).Caption, , False)
                    Btn.IconId = ImgListKey2ImgListIdx(ImgLMenu, objNode.Image)
                    If bolAddSeparator Then Btn.BeginGroup = True
                    Call CaricaSubMenuXCB(objNode.key, Btn)
                    bolAddSeparator = McolMenu(objNode.key).HasSeparator
                Else
                    Set Btn = .Add(xtpControlButton, McolMenu(objNode.key).MenuID, McolMenu(objNode.key).Caption)
                    Btn.IconId = ImgListKey2ImgListIdx(ImgLMenu, objNode.Image)
                    If bolAddSeparator Then Btn.BeginGroup = True
                    bolAddSeparator = McolMenu(objNode.key).HasSeparator
                End If
                Set objNode = objNode.Next
            Loop
        End If
    End With
        
        
End Sub



Public Function ImgListKey2ImgListIdx(ctlImgList As MSComctlLib.ImageList, ByVal ImageKey As String) As Long
    Dim i As Long
    ImgListKey2ImgListIdx = -1
    For i = 1 To ctlImgList.ListImages.Count
        If ctlImgList.ListImages(i).key = ImageKey Then
            ImgListKey2ImgListIdx = ctlImgList.ListImages(i).Tag
            Exit For
        End If
    Next i
End Function



Private Sub LoadNode(ByVal strSection As String, ByRef objParentNode As Node, ByVal bolModulo As Boolean)

Dim strKey As String
Dim strItem As Variant
Dim vetLines() As String
Dim vetValues() As String
Dim strLine As String
Dim i As Integer
Dim objNode As Node
Dim objMenu As CGestMenu

    On Local Error Resume Next
    strItem = colValues("K" & strSection)
    vetLines = Split(strItem, vbCrLf)
    For i = LBound(vetLines) To UBound(vetLines)
        strLine = Mid(vetLines(i), InStr(1, vetLines(i), "=") + 1)
        vetValues = Split(strLine, ";")
        If UBound(vetValues) > 3 Then
            If vetValues(2) <> "-" Then
                strKey = vetValues(0)
                If Len(vetValues(1)) > 0 Then
                    strKey = strKey & "_" & vetValues(1)
                End If
                On Local Error Resume Next
                strSection = colSection("K" & strKey)
                If Err.Number > 0 Then
                    strSection = ""
                End If
                On Local Error GoTo 0
                If Len(strSection) > 0 Then
                    Call ControllaKey(vetValues(3), False)
                    On Local Error Resume Next
                    Set objNode = twMenu.Nodes.Add(objParentNode, tvwChild, strKey, Replace(MXNU.CaricaCaptionInLingua(vetValues(2)), "&", ""), vetValues(3))
                    If Err.Number = 0 Then
                        If bolModulo Then
                            MstrModuloCorrente = strKey
                            'Call MostraMessaggioAccessi(9009, strSection)
                        End If
                        Call LoadNode(strSection, objNode, False)
                    End If
                    On Local Error GoTo 0
                Else
                    Call ControllaKey(vetValues(3), True)
                    Set objNode = twMenu.Nodes.Add(objParentNode, tvwChild, strKey, Replace(MXNU.CaricaCaptionInLingua(vetValues(2)), "&", ""), vetValues(3))
                End If
                    
                Set objMenu = New CGestMenu
                With objMenu
                    .Caption = MXNU.CaricaCaptionInLingua(vetValues(2))
                    .key = strKey
                    .HasSeparator = False
                    On Local Error Resume Next
                    If UBound(vetValues) = 4 Then
                        .HelpContextID = vetValues(4)
                    Else
                        .HelpContextID = vetValues(5)
                    End If
                    On Local Error GoTo 0
                    objNode.Tag = .HelpContextID    'Anomalia 6997
                    .ParentKey = objParentNode.key
                    .Modulo = MstrModuloCorrente
                    .MenuID = MlngMenuID
                End With
                McolMenu.Add objMenu
                MlngMenuID = MlngMenuID + 1
                Set objMenu = Nothing
            Else
                McolMenu(strKey).HasSeparator = True
            End If
        ElseIf UBound(vetValues) > 1 Then
            If vetValues(2) = "-" Then
                McolMenu(strKey).HasSeparator = True
            End If
        End If
    Next
    
End Sub




Private Sub ControllaKey(ImageKey As Variant, flgFoglia As Boolean)
    Dim key As Variant
    
    On Local Error Resume Next
    key = twMenu.ImageList.ListImages(ImageKey).key
    If Err <> 0 Then
        Err.Clear
        If flgFoglia Then
            ImageKey = "Exe"
        Else
            ImageKey = "CartCh"
        End If
    End If
End Sub



