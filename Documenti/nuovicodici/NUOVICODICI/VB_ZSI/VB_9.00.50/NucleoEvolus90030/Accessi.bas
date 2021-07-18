Attribute VB_Name = "MAccessi"
Option Explicit
DefLng A-Z

'===============================================
'       definizione costanti globali
'===============================================
'indici schede particolari
Global Const ID_SCHEDA_FORM = -1
Global Const ID_SCHEDA_SITUAZIONE = -2
'costanti accessi
Global Const ACC_NONDEFINITO = -1
Global Const ACC_NESSUNO = 0
Global Const ACC_LETTURA = 1
Global Const ACC_MODIFICA = 2
Global Const ACC_INSERISCI = 4
Global Const ACC_ANNULLA = 8
Global Const ACC_TUTTI = 15

Const MNU_ITEM_GESTIONE_ACCESSI = 3
Const MNU_ITEM_GESTIONE_CAMBIOUTENTE = 4

Const SEP_INDICE = "_"

'===============================================
'       definizione oggetti
'===============================================
Global myFormDefAcc As Form
Global myFormIntro As Form
Private dicEccezioni As MXKit.cDictionary
Private dicAccessi As MXKit.cDictionary
Private ControlsAccess As MXKit.cDictionary
Private AllFormControls As New Collection
'===============================================
'       definizione variabili
'===============================================
Dim mStrBufferModuliChiave As String
Dim mStrLogFile As String
Private bolRuoloTuttoAbilitato As Boolean

'======================================================================================================
'           FUNZIONI PRIVATE DEL MODULO
'======================================================================================================
Private Function AccessiFive() As Boolean
Dim strSezFive As String
    If MXNU.ModuloRegole Then
        strSezFive = MXNU.LeggiProfilo(MXAA.PathAgenti & "\FISSI\FIVE.INI", (App.EXEName), "FIVE", "")
        AccessiFive = (strSezFive <> "")
    End If
End Function

'======================================================================================================
'           FUNZIONI PUBBLICHE DEL MODULO
'======================================================================================================
Public Function Key_RivBuffetti() As Boolean
    Key_RivBuffetti = False
End Function

Public Sub SetAccessiDictionary(cAccessi As MXKit.cDictionary, CEccezioni As MXKit.cDictionary, TuttoAbilitato As Boolean)
    Set dicEccezioni = CEccezioni
    Set dicAccessi = cAccessi
    bolRuoloTuttoAbilitato = TuttoAbilitato
End Sub

Public Sub GetAccessiDictionary(cAccessi As MXKit.cDictionary, CEccezioni As MXKit.cDictionary, TuttoAbilitato As Boolean)
    Set CEccezioni = dicEccezioni
    Set cAccessi = dicAccessi
    TuttoAbilitato = bolRuoloTuttoAbilitato
End Sub


Public Sub InizializzaBufferAccessi()
    mStrBufferModuliChiave = Space$(1000)
'    '*** START DEBUG ***
'    mStrLogFile = MXNU.GetTempFile()
'    Call MXNU.ImpostaErroriSuLog(mStrLogFile, True)
'    '*** END DEBUG ***
End Sub

'Public Sub LogAccessi()
'    Call MXNU.ChiudiErroriSuLog
'    Call frmLog.MostraFileLog(mStrLogFile)
'End Sub

Public Function AccessoLettura(ByVal intAccesso) As Boolean
    AccessoLettura = ((intAccesso And ACC_LETTURA) = ACC_LETTURA)
End Function

Public Function AccessoModifica(ByVal intAccesso) As Boolean
    AccessoModifica = ((intAccesso And ACC_MODIFICA) = ACC_MODIFICA)
End Function

Public Function AccessoInserimento(ByVal intAccesso) As Boolean
    AccessoInserimento = ((intAccesso And ACC_INSERISCI) = ACC_INSERISCI)
End Function

Public Function AccessoAnnulla(ByVal intAccesso) As Boolean
    AccessoAnnulla = ((intAccesso And ACC_ANNULLA) = ACC_ANNULLA)
End Function

'leggo il numero terminale corrispondente all'utente
Public Function LeggiNumeroTerminale(vntUserID As Variant) As Integer
Dim intq As Integer
Dim strSQL As String
Dim HSS As MXKit.CRecordSet
Dim intTrm As Integer

'    If hndDBArchivi.amministratore <> "" Then
'        strSQL = "SELECT NrTerminale FROM " & hndDBArchivi.amministratore & ".TabUtenti WHERE UserID = '" & vntUserID & "'"
'    Else
        strSQL = "SELECT NrTerminale FROM TabUtenti WHERE UserID = '" & vntUserID & "'"
'    End If
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    intTrm = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "NrTerminale", 0)
    intq = MXDB.dbChiudiSS(HSS)
    
    LeggiNumeroTerminale = intTrm
End Function

Private Function FindIndexFormLing(Node As MSXML2.IXMLDOMNode, Index As Integer, isFoundIndex As Boolean) As Boolean
    Dim tempnode As MSXML2.IXMLDOMNodeList
    Dim n As MSXML2.IXMLDOMNode
    If Not Node Is Nothing Then
        Set tempnode = Node.selectNodes("ling")
        For Each n In tempnode
            If n.Attributes.getNamedItem("index").text = Index Then
                isFoundIndex = True
            Else
                Call FindIndexFormLing(n, Index, isFoundIndex)
            End If
        Next
    End If
    
    FindIndexFormLing = isFoundIndex
End Function

'NOME           : FormImpostaAccessi
'DESCRIZIONE    : legge gli accessi per la form
'PARAMETRO 1    : form di cui definire gli accessi
'PARAMETRO 2    : maschera per i bottoni della toolbar (viene modificata in base agli accessi)
'RITORNO        : maschera accessi per la form (gli accessi possono essere letti in seguito mediante le funzioni
'                 AccessiLettura,AccessiModifica,AccessiInserimento,AccessiAnnullamento)
Public Function FormGetAccessi(ByVal frmDef As Form, lngButtonMask As Long) As Integer
Dim intAccesso As Integer
Dim lngFormID As Long
Dim bolTerminaAcc As Boolean
    If (Not (frmDef Is Nothing) And MXNU.CtrlAccessi) Then
        On Local Error Resume Next
        lngFormID = frmDef.HelpContextID

        ' se non sono definiti i dizionari
        ' allora li inizializzo ricordandomi di terminarli alla fine
        If (dicEccezioni Is Nothing) Or (dicAccessi Is Nothing) Then
            bolTerminaAcc = True
            'RIF.A#7641 - ottimizzazione caricamento accessi form
            Call InizializzaLetturaAccessi(MXNU.UtenteAttivo, False, lngFormID)
        End If
        
        'leggo l'accesso per la form...
        intAccesso = LeggiAccessi(MXNU.UtenteAttivo, lngFormID, ID_SCHEDA_FORM, False)
        If (intAccesso = ACC_NESSUNO) Then
            lngButtonMask = 0
        Else
            If (intAccesso And ACC_INSERISCI) = 0 Then lngButtonMask = lngButtonMask And (Not BTN_INS_MASK)
            If (intAccesso And ACC_MODIFICA) = 0 Then lngButtonMask = lngButtonMask And (Not BTN_REG_MASK)
            If (intAccesso And ACC_ANNULLA) = 0 Then lngButtonMask = lngButtonMask And (Not BTN_ANN_MASK)
        End If
    Else
        intAccesso = ACC_TUTTI
    End If
    
    If bolTerminaAcc Then
        Call TerminaLetturaAccessi
    End If
    
    FormGetAccessi = intAccesso
End Function

'NOME           : FormImpostaAccessi
'DESCRIZIONE    : legge ed imposta gli accessi per la form
'PARAMETRO 1    : form di cui definire gli accessi
'PARAMETRO 2    : maschera per i bottoni della toolbar (viene modificata in base agli accessi)
'RITORNO        : maschera accessi per la form (gli accessi possono essere letti in seguito mediante le funzioni
'                 AccessiLettura,AccessiModifica,AccessiInserimento,AccessiAnnullamento)
Public Function FormImpostaAccessi(ByVal frmDef As Form, lngButtonMask As Long) As Integer
Dim lngFormID As Long
Dim ctrGen As Control, ctrGen1 As Control, ctrlScheda As Control
Dim intAccesso As Integer
Dim intAccLing As Integer
Dim objContainer As Object
Dim gestisciling As Boolean
Dim schedaparent As MWSchedaBox
Dim bolTerminaAcc As Boolean

    gestisciling = True
    If (Not (frmDef Is Nothing) And MXNU.CtrlAccessi) Then
        
        
        On Local Error Resume Next
        lngFormID = frmDef.HelpContextID
        
        If Not ControlsAccess Is Nothing Then
            ControlsAccess.Clear
            Set ControlsAccess = Nothing
        End If
        Set ControlsAccess = New cDictionary
  
        'leggo l'accesso per la form...
        intAccesso = LeggiAccessi(MXNU.UtenteAttivo, lngFormID, ID_SCHEDA_FORM, False)
        If (intAccesso = ACC_NESSUNO) Then
            lngButtonMask = 0
        Else
            If (intAccesso And ACC_INSERISCI) = 0 Then lngButtonMask = lngButtonMask And (Not BTN_INS_MASK)
            If (intAccesso And ACC_MODIFICA) = 0 Then lngButtonMask = lngButtonMask And (Not BTN_REG_MASK)
            If (intAccesso And ACC_ANNULLA) = 0 Then lngButtonMask = lngButtonMask And (Not BTN_ANN_MASK)
        End If
        
        'se non trovo la form nella definizione delle linguette allora esco da qui
        Dim strPathFileFormLing As String
        Dim xmlFormLing As MSXML2.DOMDocument
        Dim xmlNode As MSXML2.IXMLDOMNode
        Dim bolFormLingCaricato As Boolean
        
        'modifiche effettuate da RZ in data 20/12 per gestire alcune anomalie sulla gestione
        'degli accessi. Ho tolto la gestione per le form di estensione perchè ho generalizzato
        'il tutto con 2 sub DisattivaWrapper e Disattivacontrolli. L'anomiala #8262 è ora
        'definitivamente risolta ed è possibile gestire accessi a linguette sottolinguette sia di form
        'standard che di form estensioni sia frmexchild che estensioni all'interno di form
        
        'prima lo cerco nel pers ditta poi in parte server metodo
        strPathFileFormLing = CercaDirFile("FormLing.xml", MXNU.PercorsoPers$ & "\" & MXNU.DittaAttiva)
        If StrComp(strPathFileFormLing, "", vbTextCompare) = 0 Then
            strPathFileFormLing = CercaDirFile("FormLing.xml", MXNU.PercorsoPgm & "\Tools")
        End If
        If StrComp(strPathFileFormLing, "", vbTextCompare) <> 0 Then
            Set xmlFormLing = New MSXML2.DOMDocument
            bolFormLingCaricato = xmlFormLing.Load(strPathFileFormLing)
            Set xmlNode = xmlFormLing.selectSingleNode("//rootNode/node[@helpid='" & lngFormID & "']")
            gestisciling = Not xmlNode Is Nothing
        End If
        If Not gestisciling Then
            FormImpostaAccessi = intAccesso
            If intAccesso = ACC_LETTURA Then
                Call DisabilitaControlli(frmDef, True)
            End If
        Else
            ' Rif. anomalia #8262 - gestione accessi per le estensioni
                '...e per le linguette
                'carico i controlli di tipo MWLinguetta e Ling
                Call LoadRecursiveControls(frmDef)
                
                For Each ctrGen In AllFormControls
                    If (TypeName(ctrGen) = "MWLinguetta" And (UCase$(ctrGen.NAME) = "LING" Or ctrGen.NAME = "LingIva")) Then  'rif.sch. A6457
                        'Gestione caso particolare per Prima Nota vedi frmDefAcc e Anomalia 3639
                        If UCase$(ctrGen.NAME) = "LING" Then   'rif.sch. A6457
                            intAccLing = LeggiAccessi(MXNU.UtenteAttivo, lngFormID, ctrGen.Index, False)
                        Else
                            intAccLing = LeggiAccessi(MXNU.UtenteAttivo, lngFormID, 9990 + ctrGen.Index, False)
                        End If
                        Set schedaparent = ctrGen.Parent.Controls("scheda")(ctrGen.Index)
                        If (intAccLing = ACC_NESSUNO) And (FindIndexFormLing(xmlNode, ctrGen.Index, False)) Then
                            'nessun accesso -> rendo invisibile la linguetta e la scheda
                            If ctrGen.NAME <> "LingIva" Then
                                Call ControlsAccess.Add("#" & ctrGen.hwnd, True, "#" & ctrGen.hwnd)
                                ctrGen.Enabled = False
                                schedaparent.Visible = False
                                If Not schedaparent Is Nothing Then Call DisabilitaControlli(schedaparent, True, False)
                            Else
                                frmDef.LingIva(ctrGen.Index).Enabled = False
                                frmDef.SchIva(ctrGen.Index).Enabled = False
                                For Each ctrGen1 In frmDef.SchIva(ctrGen.Index).Controls
                                    ctrGen1.Visible = False
                                Next ctrGen1
                            End If
                        ElseIf (intAccLing = ACC_LETTURA) Then
                            'accesso sola lettura -> abilito la linguetta e disabilito la scheda
                            If ctrGen.NAME <> "LingIva" Then
                                If ControlsAccess.Exists("#" & ctrGen.hwnd) Then 'l'ho precedentemente disabilitata
                                    ctrGen.Enabled = True
                                    schedaparent.Visible = True
                                End If
                                ' Rif .anomalia #6911
                                If (schedaparent.Controls.Count = 0) And (schedaparent.ControlsEx.Count = 0) Then
                                    schedaparent.Enabled = False
                                Else
                                    'disabilito i controlli
                                    Call DisabilitaControlli(schedaparent, True)
                                End If
                                schedaparent.TabStop = True
                                schedaparent.TabIndex = ctrGen.TabIndex + 1
                            Else
                                frmDef.LingIva(ctrGen.Index).Enabled = True
                                For Each ctrGen1 In frmDef.SchIva(ctrGen.Index).Controls
                                    'If Not (TypeName(ctrGen1) = "MWLinguetta" And ctrGen1.Name = "Ling") Then  'Rif. Anomalie98 1533
                                        ctrGen1.Enabled = False
                                        If TypeName(ctrGen1) = "fpSpread" Then
                                            ctrGen1.Tag = "DONTENABLE"
                                        End If
                                    'End If
                                Next ctrGen1
                                frmDef.SchIva(ctrGen.Index).TabStop = True
                                frmDef.SchIva(ctrGen.Index).TabIndex = frmDef.LingIva(ctrGen.Index).TabIndex + 1
                            End If
                        '#11857 per la gestione accessi delle sottolinguette caso in cui metto in sola lettura una
                        'lingueta principale e voglio poi mettere in scritture le sottolinguette
                        ElseIf (intAccLing >= ACC_MODIFICA) Then
                            'accesso in scrittura-> abilito la linguetta e i controlli
                                If ctrGen.NAME <> "LingIva" Then
                                If ControlsAccess.Exists("#" & ctrGen.hwnd) Then 'l'ho precedentemente disabilitata
                                    ctrGen.Enabled = True
                                    schedaparent.Visible = True
                                End If
                                ' Rif .anomalia #6911
                                If (schedaparent.Controls.Count = 0) And (schedaparent.ControlsEx.Count = 0) Then
                                    schedaparent.Enabled = True
                                Else
                                    Call DisabilitaControlli(schedaparent, False)
                                End If
                            Else
                                frmDef.LingIva(ctrGen.Index).Enabled = True
                                'frmDef.Scheda(ctrGen.Index).Enabled = False
                                For Each ctrGen1 In frmDef.SchIva(ctrGen.Index).Controls
                                    'If Not (TypeName(ctrGen1) = "MWLinguetta" And ctrGen1.Name = "Ling") Then  'Rif. Anomalie98 1533
                                        ctrGen1.Enabled = False
                                        If TypeName(ctrGen1) = "fpSpread" Then
                                            ctrGen1.Tag = "DONTENABLE"
                                        End If
                                    'End If
                                Next ctrGen1
                            End If
                        
                        End If
                    End If
                Next
            'End If
        End If
    Else
        intAccesso = ACC_TUTTI
    End If
    
    '------BEGIN distruggo gli oggetti cachati---------
    If Not AllFormControls Is Nothing Then
        Dim i As Integer
        For i = AllFormControls.Count To 1 Step -1
            Call AllFormControls.Remove(i)
        Next
        Set AllFormControls = Nothing
    End If
    '-----------------END---------------
    
    If bolTerminaAcc Then
        Call TerminaLetturaAccessi
    End If
    
    FormImpostaAccessi = intAccesso
End Function

Private Sub DisabilitaWrapper(objContainer As Control, Disable As Boolean, Optional Visible As Boolean = True)
Dim ctrGen1 As Control

    On Local Error Resume Next
    For Each ctrGen1 In objContainer.object.Controls(1).Controls
        If Not (TypeName(ctrGen1) = "MWLinguetta" And (UCase$(ctrGen1.NAME) = "LING" Or ctrGen1.NAME = "LingIva" Or UCase$(ctrGen1.NAME) = "SUBLING")) Then    'Rif. Anomalie98 1533   'rif.sch. A6457
            If TypeName(ctrGen1) <> "MWSchedaBox" Then
                If Disable Then 'sto disabilitando il controllo
                    If ctrGen1.Enabled Then
                        ControlsAccess.Add "#" & ctrGen1.hwnd, Disable, "#" & ctrGen1.hwnd
                        If TypeName(ctrGen1) = "fpSpread" Then
                            If ctrGen1.Tag <> "DONTENABLE" Then
                                ctrGen1.Tag = "DONTENABLE"
                                ctrGen1.OperationMode = 1 ' SS_OP_MODE_READONLY
                                ctrGen1.Enabled = True
                            End If
                        Else
                            ctrGen1.Enabled = Not Disable
                        End If
                    End If
                Else
                    'riabilito solamente i controlli che ho precedentemente disabilitato
                   If ControlsAccess.Exists("#" & ctrGen1.hwnd) Then
                        If TypeName(ctrGen1) = "fpSpread" Then
                            If ctrGen1.Tag = "DONTENABLE" Then
                                ctrGen1.Tag = ""
                                ctrGen1.OperationMode = 0 ' SS_OP_MODE_READ_WRITE
                                ctrGen1.Enabled = True
                            End If
                        Else
                            ctrGen1.Enabled = Not Disable
                        End If
                   End If
                End If
            Else
                ctrGen1.TabStop = True
            End If
        End If
        ctrGen1.Visible = Visible
    Next
End Sub
Private Sub LoadWrapperControls(objContainer As Object)
On Local Error Resume Next

    Dim ctrGen1 As Control
    Dim swapctrl As Control

    
    For Each ctrGen1 In objContainer.object.Controls(1).Controls
        If (TypeName(ctrGen1) = "MWLinguetta" And (UCase$(ctrGen1.NAME) = "LING" Or ctrGen1.NAME = "LingIva")) Then     'Rif. Anomalie98 1533   'rif.sch. A6457
            Set swapctrl = AllFormControls("K" & ctrGen1.hwnd)
            If Err = 5 Then
                 Call AllFormControls.Add(ctrGen1, "K" & ctrGen1.hwnd)
            End If
            On Local Error Resume Next
        End If
    Next

End Sub

Private Sub LoadRecursiveControls(objContainer As Object)
On Local Error Resume Next

    Dim ctrGen1 As Control
    Dim ctlrs As Integer
    ctlrs = 0
    
    Dim swapctrl As Control
    
    If TypeOf objContainer Is Form Then
        For Each ctrGen1 In objContainer.Controls
            If Not ctrGen1 Is Nothing Then
                If LCase(ctrGen1.NAME) = "objextwrapper" Then
                    Call LoadWrapperControls(ctrGen1)
                End If
                If (TypeName(ctrGen1) = "MWLinguetta" And (UCase$(ctrGen1.NAME) = "LING" Or ctrGen1.NAME = "LingIva")) Then
                    Set swapctrl = AllFormControls("K" & ctrGen1.hwnd)
                    If Err = 5 Then
                        Call AllFormControls.Add(ctrGen1, "K" & ctrGen1.hwnd)
                    End If
                    On Local Error Resume Next
                Else
                    Call LoadRecursiveControls(ctrGen1)
                End If
            End If
        Next
    ElseIf TypeOf objContainer Is Frame Then
        For Each ctrGen1 In objContainer.Parent.Controls
            If ctrGen1.Container Is objContainer Then
                If LCase(ctrGen1.NAME) = "objextwrapper" Then
                    Call LoadWrapperControls(ctrGen1)
                End If
                If (TypeName(ctrGen1) = "MWLinguetta" And (UCase$(ctrGen1.NAME) = "LING" Or ctrGen1.NAME = "LingIva")) Then
                    Set swapctrl = AllFormControls(ctrGen1.hwnd)
                    If Err = 5 Then
                         Call AllFormControls.Add(ctrGen1, "K" & ctrGen1.hwnd)
                    End If
                    On Local Error Resume Next
                Else
                    Call LoadRecursiveControls(ctrGen1)
                End If
            End If
        Next
    Else
        For Each ctrGen1 In objContainer.object.Controls
            If Not ctrGen1 Is Nothing Then
                If LCase(ctrGen1.NAME) = "objextwrapper" Then
                    Call LoadWrapperControls(ctrGen1)
                End If
                If (TypeName(ctrGen1) = "MWLinguetta" And (UCase$(ctrGen1.NAME) = "LING" Or ctrGen1.NAME = "LingIva")) Then
                    Set swapctrl = AllFormControls(ctrGen1.hwnd)
                    If Err = 5 Then
                         Call AllFormControls.Add(ctrGen1, "K" & ctrGen1.hwnd)
                    End If
                    On Local Error Resume Next
                Else
                    Call LoadRecursiveControls(ctrGen1)
                End If
            End If
        Next
        For Each ctrGen1 In objContainer.object.ControlsEx
            If Not ctrGen1 Is Nothing Then
                If LCase(ctrGen1.NAME) = "objextwrapper" Then
                    Call LoadWrapperControls(ctrGen1)
                End If
                If (TypeName(ctrGen1) = "MWLinguetta" And (UCase$(ctrGen1.NAME) = "LING" Or ctrGen1.NAME = "LingIva")) Then
                   Set swapctrl = AllFormControls(ctrGen1.hwnd)
                    If Err = 5 Then
                         Call AllFormControls.Add(ctrGen1, "K" & ctrGen1.hwnd)
                    End If
                    On Local Error Resume Next
                Else
                    Call LoadRecursiveControls(ctrGen1)
                End If
            End If
        Next
    End If
    
End Sub

Private Sub DisabilitaControlli(objContainer As Object, Disable As Boolean, Optional Visible As Boolean = True)
Dim mycontrols As New Collection
    On Local Error Resume Next
    Dim ctrGen1 As Control
    Dim ctlrs As Integer
    ctlrs = 0
    
    
    If TypeOf objContainer Is Form Then
        For Each ctrGen1 In objContainer.Controls
            If Not ctrGen1 Is Nothing Then
                mycontrols.Add ctrGen1
            End If
        Next
    ElseIf TypeOf objContainer Is Frame Then
        For Each ctrGen1 In objContainer.Parent.Controls
            If ctrGen1.Container Is objContainer Then
                mycontrols.Add ctrGen1
            End If
        Next
    Else
        For Each ctrGen1 In objContainer.object.Controls
            If Not ctrGen1 Is Nothing Then
                mycontrols.Add ctrGen1
            End If
        Next
    
        For Each ctrGen1 In objContainer.object.ControlsEx
            If Not ctrGen1 Is Nothing Then
                mycontrols.Add ctrGen1
            End If
        Next
    End If
    
    'in alcune parti di metodo viene controllata la proprietà enabled e non visibile per impostare proprietà di controlli dipendenti
    'enabled diventa false quando la proprietà visible è a false quando l'accesso è disabilitato. se imposto la proprietà enabled a false
    'per frame e mwschedabox viene disabilitata automagicamente la linguetta
    If Not ((TypeName(objContainer) = "MWLinguetta") And (UCase$(objContainer.NAME) = "LING") Or (UCase$(objContainer.NAME) = "SUBLING") Or (UCase$(objContainer.NAME) = "SUBLINGTP") Or (UCase$(objContainer.NAME) = "SUBLINGTPGEN")) Then    'Rif. Anomalie98 1533   'rif.sch. A6457
       If Not Visible Then
            If Not ControlsAccess.Exists(objContainer.hwnd) Then
                ControlsAccess.Add "#" & objContainer.hwnd, Disable, "#" & objContainer.hwnd
            End If
            objContainer.Visible = Visible
       Else
          If ControlsAccess.Exists("#" & objContainer.hwnd) Then
            objContainer.Visible = Visible
          End If
       End If
    End If
    
    For Each ctrGen1 In mycontrols
        If LCase(ctrGen1.NAME) = "objextwrapper" Then
            Call DisabilitaWrapper(ctrGen1, Disable, Visible)
        ElseIf Not ((TypeName(ctrGen1) = "MWLinguetta") And (UCase$(ctrGen1.NAME) = "LING") Or (UCase$(ctrGen1.NAME) = "SUBLING") Or (UCase$(ctrGen1.NAME) = "SUBLINGTP") Or (UCase$(ctrGen1.NAME) = "SUBLINGTPGEN")) Then    'Rif. Anomalie98 1533   'rif.sch. A6457
            If TypeName(ctrGen1) <> "MWSchedaBox" And TypeName(ctrGen1) <> "Frame" Then
                If Disable Then 'sto disabilitando il controllo
                    If ctrGen1.Enabled Then
                        ControlsAccess.Add "#" & ctrGen1.hwnd, Disable, "#" & ctrGen1.hwnd
                        If TypeName(ctrGen1) = "fpSpread" Then
                            If ctrGen1.Tag <> "DONTENABLE" Then
                                ctrGen1.Tag = "DONTENABLE"
                                ctrGen1.OperationMode = 1 ' SS_OP_MODE_READONLY
                                ctrGen1.Enabled = True
                            End If
                        Else
                            If TypeOf ctrGen1 Is TextBox Then
                                If ctrGen1.MultiLine Then
                                    ctrGen1.Locked = Disable
                                Else
                                    ctrGen1.Enabled = Not Disable
                                End If
                            Else
                                ctrGen1.Enabled = Not Disable
                            End If
                        End If
                        If Not ctrGen1.Enabled Then
                            If ctrGen1.Tag = "DONTDISABLE" Then
                                ctrGen1.Enabled = True
                            End If
                        End If
                    End If
                Else
                   'riabilito solamente i controlli che ho precedentemente disabilitato
                   If ControlsAccess.Exists("#" & ctrGen1.hwnd) Then
                        If TypeName(ctrGen1) = "fpSpread" Then
                            If ctrGen1.Tag = "DONTENABLE" Then
                                ctrGen1.Tag = ""
                                ctrGen1.OperationMode = 0 ' SS_OP_MODE_READ_WRITE
                                ctrGen1.Enabled = True
                            End If
                        Else
                            If TypeOf ctrGen1 Is TextBox Then
                                If ctrGen1.MultiLine Then
                                    ctrGen1.Locked = Disable
                                Else
                                    ctrGen1.Enabled = Not Disable
                                End If
                            Else
                                ctrGen1.Enabled = Not Disable
                            End If
                        End If
                   End If
                End If
            Else
                DisabilitaControlli ctrGen1, Disable, Visible
                ctrGen1.TabStop = True
            End If
        End If
    Next
End Sub

Function MenuDefinisciAccessi(mnuGen As Object) As Boolean    'mnuGen As MXKit.CMenu

Dim vntNome As Variant
Dim vntIndex As Variant
Dim bolCtrlAccessi As Boolean

    On Local Error Resume Next
    vntNome = "mnu" & ScomponiChiave(mnuGen.key, vntIndex)
    If (Err <> 0) Then vntIndex = ""
    On Local Error GoTo 0
    
    '[06/05/2011] Rimozione Chiave Hardware
    'MenuDefinisciAccessi = ChiaveDammiAccessi(vntNome, vntIndex)
    MenuDefinisciAccessi = MXNU.AssegnaModuliChiave(vntNome, vntIndex)
    If (MenuDefinisciAccessi) Then
        If (MXNU.CtrlAccessi) Then
            MenuDefinisciAccessi = MenuDefinisciAccessiSupervisor(vntNome, vntIndex, bolCtrlAccessi)
            If MenuDefinisciAccessi And bolCtrlAccessi And mnuGen.HelpContextID <> 0 Then
                MenuDefinisciAccessi = (LeggiAccessi(MXNU.UtenteAttivo, mnuGen.HelpContextID, ID_SCHEDA_FORM, True) > ACC_NESSUNO)
            End If
        End If
    End If
    
    'Correzione anomalia (l'utente non amministratore non appartenente a nessun gruppo non vedeva il menu MENU)     ...e non poteva più uscire da metodo
    If InStr(1, vntNome, "mnuMenu", vbTextCompare) > 0 Then MenuDefinisciAccessi = True
    
    'If Len(vntIndex) > 0 Then mnuGen.Visible = MenuDefinisciAccessi
    'mnuGen.Enabled = MenuDefinisciAccessi
End Function

Function KeyGruppoGet(vntGruppo As Variant) As Variant
    KeyGruppoGet = "gruppo" & vntGruppo
End Function

Function KeyUtenteGet(vntGruppo As Variant, vntUtente As Variant) As Variant
    KeyUtenteGet = vntGruppo & "_" & vntUtente
End Function

'RIF.A#7641 - se lettura menu => IndiceScheda è sempre -1
'             se lettura form => filtro per HelpId e non filtro la query per IndiceScheda
Sub InizializzaLetturaAccessi(ByVal strUtente As String, ByVal bolLetturaMenu As Boolean, Optional ByVal lngHelpID As Long = 0)
    On Local Error GoTo InizializzaLetturaAccessi_ERR
    
    Dim strSQL As String
    Dim sLngHelpID As Long
    Dim sIntIndiceScheda As Integer
    Dim sIntTipoAccesso As Integer
    Dim strKey As String
    Dim i As Integer
    Dim bolUtenteInGruppo As Boolean
    
    Dim HSS As MXKit.CRecordSet
    Dim dicTemp As MXKit.cDictionary
    Dim dtAcc As MXKit.cDatiAccesso
    
    bolUtenteInGruppo = False
    bolRuoloTuttoAbilitato = False

    ' controllo se l'utente fa parte di almeno un gruppo
    strSQL = "SELECT CodGruppo" & _
            " FROM TabMembriGruppo" & _
            " WHERE CodUtente=" & hndDBArchivi.FormatoSQL(strUtente, DB_TEXT)
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    bolUtenteInGruppo = Not MXDB.dbFineTab(HSS)
    If Not HSS Is Nothing Then
        MXDB.dbChiudiSS HSS
        Set HSS = Nothing
    End If
    
    ' Il funzionamento della nuova gestione accessi prevede che un utente
    '    può avere uno o più profili (quindi un ruolo associato)
    ' OPPURE
    '    può appartenere ad uno o più gruppi (i quali a loro volta avranno dei profili)
    ' Per questo se l'utente appartiene ad almeno un gruppo non controllo l'esistenza
    ' di ruoli associati direttamente all'utente
    ' In ogni caso se non trovo ruoli associati è come se fosse impostato di default il ruolo 0 (Tutto Disabilitato)
    
    Set dicEccezioni = New MXKit.cDictionary
    Set dicAccessi = New MXKit.cDictionary

    ' inserisco nel dizionario prima di tutto le eccezioni dell'Utente perchè hanno la priorità
    strSQL = "SELECT HelpID, IndiceScheda, TipoAccesso" & _
            " FROM TabAccessiUtente" & _
            " WHERE CodUtente=" & hndDBArchivi.FormatoSQL(strUtente, DB_TEXT)
    'RIF.A#7641 - se HelpID = -1 => lettura accessi menu
    If (bolLetturaMenu) Then
        strSQL = strSQL '& " AND IndiceScheda=-1"
    Else
        strSQL = strSQL & " AND HelpId=" & lngHelpID
    End If
    GoSub CaricaDaQueryInTemp
    GoSub CaricaEccezDaTemp
    
    ' inserisco nel dizionario le eccezioni dei Gruppi
    ' solo se l'Utente appartiene ad almeno un Gruppo
    If bolUtenteInGruppo Then
        ' se l'eccezione è già presente nel dizionario, va lasciata quella dell'Utente
        ' in caso di conflitto fra eccezioni dei Gruppi vince l'eccezione con PIU' permessi
        strSQL = "SELECT AG.HelpID, AG.IndiceScheda, AG.TipoAccesso" & _
                " FROM TabAccessiGruppo AG INNER JOIN TabMembriGruppo MG" & _
                " ON AG.CodGruppo=MG.CodGruppo" & _
                " WHERE MG.CodUtente=" & hndDBArchivi.FormatoSQL(strUtente, DB_TEXT)
        'RIF.A#7641 - se HelpID = -1 => lettura accessi menu
        If (bolLetturaMenu) Then
            strSQL = strSQL '& " AND AG.IndiceScheda=-1"
        Else
            strSQL = strSQL & " AND AG.HelpId=" & lngHelpID
        End If
    End If
    GoSub CaricaDaQueryInTemp
    GoSub CaricaEccezDaTemp
    
    ' inserisco nel dizionario gli accessi dei Ruoli
    ' CASO PARTICOLARE
    ' controllo se è presente un CodRuolo=-1
    strSQL = "SELECT P.Progressivo" & _
        " FROM TabProfili P"
    If bolUtenteInGruppo Then
        strSQL = strSQL & _
        " INNER JOIN TabMembriGruppo MG" & _
        " ON P.CodGruppo=MG.CodGruppo" & _
        " WHERE MG."
    Else
        strSQL = strSQL & _
        " WHERE P."
    End If
    strSQL = strSQL & "CodUtente=" & hndDBArchivi.FormatoSQL(strUtente, DB_TEXT) & " AND P.CodRuolo=-1"
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    bolRuoloTuttoAbilitato = Not MXDB.dbFineTab(HSS)
    If Not HSS Is Nothing Then
        MXDB.dbChiudiSS HSS
        Set HSS = Nothing
    End If
        
    If Not bolRuoloTuttoAbilitato Then
        ' se l'accesso è già presente nel dizionario, va lasciato (perchè eccezione dei Gruppi o dell'Utente)
        ' in caso di conflitto fra accessi dei Ruoli, vince l'accesso con PIU' permessi
        strSQL = "SELECT AR.HelpID, AR.IndiceScheda, AR.TipoAccesso" & _
            " FROM TabProfili P INNER JOIN TabAccessiRuolo AR" & _
            " ON P.CodRuolo = AR.CodRuolo"
        If bolUtenteInGruppo Then
            strSQL = strSQL & _
            " INNER JOIN TabMembriGruppo MG" & _
            " ON P.CodGruppo=MG.CodGruppo" & _
            " WHERE MG."
        Else
            strSQL = strSQL & _
            " WHERE P."
        End If
        strSQL = strSQL & "CodUtente=" & hndDBArchivi.FormatoSQL(strUtente, DB_TEXT)
        'RIF.A#7641 - se HelpID = -1 => lettura accessi menu
        If (bolLetturaMenu) Then
            strSQL = strSQL '& " AND AR.IndiceScheda=-1"
        Else
            strSQL = strSQL & " AND AR.HelpId=" & lngHelpID
        End If
        GoSub CaricaDaQueryInTemp
        GoSub CaricaAccessiDaTemp
    End If
    
InizializzaLetturaAccessi_END:
    On Local Error GoTo 0
    If Not HSS Is Nothing Then
        MXDB.dbChiudiSS HSS
        Set HSS = Nothing
    End If
    Set dicTemp = Nothing
    Set dtAcc = Nothing
    Exit Sub

InizializzaLetturaAccessi_ERR:
    Dim lngErrCod As Long
    Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("[MetodoEvolus].[MAccessi].[InizializzaLetturaAccessi]", lngErrCod, strErrDsc))
    Resume InizializzaLetturaAccessi_END
    Resume
    
CaricaDaQueryInTemp:
    Set dicTemp = New MXKit.cDictionary
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    While Not MXDB.dbFineTab(HSS, TIPO_SNAPSHOT)
        ' mi salvo i dati della riga
        sLngHelpID = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "HELPID", 0)
        sIntIndiceScheda = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "INDICESCHEDA", -1)
        sIntTipoAccesso = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "TIPOACCESSO", -1)
        strKey = CStr(sLngHelpID) & ";" & CStr(sIntIndiceScheda)
        If dicTemp.Exists(strKey) Then
            ' se esiste modifico
            Set dtAcc = dicTemp(strKey).value
            ' devo tenere PIU' accessi possibile, per questo faccio un OR fra gli accessi
            ' CONTROLLO per sicurezza che l'accesso sia valido
            If (sIntTipoAccesso > ACC_NONDEFINITO) And (sIntTipoAccesso <= ACC_TUTTI) Then
                dtAcc.TipoAccesso = dtAcc.TipoAccesso Or sIntTipoAccesso
            End If
        Else
            ' se non esiste aggiungo
            Set dtAcc = New MXKit.cDatiAccesso
            dtAcc.HelpID = sLngHelpID
            dtAcc.IndiceScheda = sIntIndiceScheda
            dtAcc.TipoAccesso = sIntTipoAccesso
            dicTemp.Add strKey, dtAcc
        End If
        Set dtAcc = Nothing
        MXDB.dbSuccessivo HSS, TIPO_SNAPSHOT
    Wend
    If Not HSS Is Nothing Then
        MXDB.dbChiudiSS HSS
        Set HSS = Nothing
    End If
    ' ho messo in dicTemp tutti gli accessi o tutte le eccezioni correnti
    Return
    
CaricaEccezDaTemp:
    ' ora copio in dicEccezioni i dati raccolti nel dicTemp scartando i dati già presenti
    For i = 1 To dicTemp.Count
        If Not dicEccezioni.Exists(dicTemp(i).NAME) Then
            dicEccezioni.Add dicTemp(i).NAME, dicTemp(i).value
        End If
    Next i
    Set dicTemp = Nothing
    Return
    
CaricaAccessiDaTemp:
    ' ora copio in dicAccessi i dati raccolti nel dicTemp scartando i dati già presenti
    For i = 1 To dicTemp.Count
        If Not dicAccessi.Exists(dicTemp(i).NAME) Then
            dicAccessi.Add dicTemp(i).NAME, dicTemp(i).value
        End If
    Next i
    Set dicTemp = Nothing
    Return
End Sub

Sub TerminaLetturaAccessi()
    If Not dicEccezioni Is Nothing Then
        dicEccezioni.Clear
        Set dicEccezioni = Nothing
    End If
    If Not dicAccessi Is Nothing Then
        dicAccessi.Clear
        Set dicAccessi = Nothing
    End If
End Sub

'Utente appartenente ad un gruppo
'GRUPPO     D       D       N       N
'UTENTE     D       N       N       D
'ACCESSI    U       G      ALL      U
'           1a      2a      3a      4a
'Utente NON appartenente a gruppi
'UTENTE     D       N
'ACCESSI    U       NO
'           1b      2b
Function LeggiAccessi(ByVal strUtente As String, ByVal lngFormID As Long, ByVal intScheda As Integer, Optional ByVal bolLetturaMenu As Boolean = False) As Integer
    On Local Error GoTo LeggiAccessi_ERR
    
    Dim intLeggiAccessi As Integer
    Dim bolTerminaAcc As Boolean
    Dim strKey As String
    
    Dim dtAcc As MXKit.cDatiAccesso
    
    bolTerminaAcc = False
    intLeggiAccessi = ACC_NESSUNO
    
    ' se non sono definiti i dizionari
    ' allora li inizializzo ricordandomi di terminarli alla fine
    If (dicEccezioni Is Nothing) Or (dicAccessi Is Nothing) Then
        bolTerminaAcc = True
        Call InizializzaLetturaAccessi(strUtente, bolLetturaMenu, lngFormID)
    End If
    
    strKey = CStr(lngFormID) & ";" & IIf(bolLetturaMenu, "-1", CStr(intScheda))
    ' controllo se è presente un'eccezione
    If dicEccezioni.Exists(strKey) Then
        Set dtAcc = dicEccezioni(strKey).value
        intLeggiAccessi = dtAcc.TipoAccesso
    Else
        ' se non è presente un'eccezione
        ' CASO PARTICOLARE
        ' controllo se l'utente eredita dal ruolo Tutto Abilitato
        If bolRuoloTuttoAbilitato Then
            intLeggiAccessi = ACC_TUTTI
        Else
            'controllo se è presente un accesso fra i ruoli
            If dicAccessi.Exists(strKey) Then
                Set dtAcc = dicAccessi(strKey).value
                intLeggiAccessi = dtAcc.TipoAccesso
            End If
        End If
    End If
    
    'nel caso in cui l'accesso non è definito, ritorno nessun accesso
    If (intLeggiAccessi = ACC_NONDEFINITO) Then intLeggiAccessi = ACC_NESSUNO
    
    If bolTerminaAcc Then
        Call TerminaLetturaAccessi
    End If
    
LeggiAccessi_END:
    On Local Error GoTo 0
    LeggiAccessi = intLeggiAccessi
    Set dtAcc = Nothing
    Exit Function

LeggiAccessi_ERR:
    Dim lngErrCod As Long
    Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    intLeggiAccessi = ACC_NESSUNO
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("[MetodoEvolus].[MAccessi].[LeggiAccessi]", lngErrCod, strErrDsc))
    Resume LeggiAccessi_END
    Resume
End Function

Private Function MenuDefinisciAccessiSupervisor(vntNome As Variant, vntIndex As Variant, bolCtrlAccessi As Boolean) As Boolean

    Select Case vntNome
        Case "mnuStrumItem"
            If (vntIndex = MNU_ITEM_GESTIONE_ACCESSI) Then
                MenuDefinisciAccessiSupervisor = (Not MXNU.CtrlAccessi)
                bolCtrlAccessi = False
            ElseIf (vntIndex = MNU_ITEM_GESTIONE_CAMBIOUTENTE) Then
                MenuDefinisciAccessiSupervisor = True
                bolCtrlAccessi = False
            Else
                'rif.sch.2881 - agli altri menu degli strumenti deve essere dato accesso anche ad utente non supervisor
                MenuDefinisciAccessiSupervisor = True
                bolCtrlAccessi = True
            End If
        Case "mnuStrumItem_" & MNU_ITEM_GESTIONE_ACCESSI
            MenuDefinisciAccessiSupervisor = (Not MXNU.CtrlAccessi)
            bolCtrlAccessi = False
        Case "mnuStrumItem_" & MNU_ITEM_GESTIONE_CAMBIOUTENTE
            MenuDefinisciAccessiSupervisor = True
            bolCtrlAccessi = False
        Case Else
            MenuDefinisciAccessiSupervisor = True
            bolCtrlAccessi = True
    End Select
    
End Function

Sub TreeUtenti_Inizializza(trwUtenti As TreeView, bolGruppoUtenti As Boolean, ByVal bolIncludiSupervisor As Boolean)
    
    Dim intq As Integer
    Dim nodX As Node
    Dim HSS As CRecordSet
    Dim strSQL As String
    Dim bolEnd As Boolean
    Dim vntCodGrp As Variant
    Dim vntDscGrp As Variant
    Dim vntCodUte As Variant

    'imposto radice
    Set nodX = trwUtenti.Nodes.Add(, tvwFirst, "Metodo98", "Metodo", "metodo98")
    Call nodX.EnsureVisible
    'imposto gruppi
    Set nodX = trwUtenti.Nodes.Add("Metodo98", tvwChild, "gruppi", "Gruppi", "entire")
    Call nodX.EnsureVisible
    strSQL = "SELECT Codice,Descrizione FROM TabGruppiUtente ORDER BY Codice"
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    bolEnd = MXDB.dbFineTab(HSS, TIPO_SNAPSHOT)
    
    If bolIncludiSupervisor Then
        strSQL = "SELECT UserID" _
                & " FROM {oj TabUtenti TU LEFT OUTER JOIN TabMembriGruppo MG ON MG.CodUtente=TU.UserID}" _
                & " WHERE CodGruppo="
    Else
        strSQL = "SELECT UserID" _
                & " FROM {oj TabUtenti TU LEFT OUTER JOIN TabMembriGruppo MG ON MG.CodUtente=TU.UserID}" _
                & " WHERE Supervisor=0 AND CodGruppo="
    End If
    
    Do While (Not bolEnd)
        vntCodGrp = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "Codice", 0)
        If (vntCodGrp <> 0) Then
            vntDscGrp = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "Descrizione", "")
            If (vntDscGrp = "") Then vntDscGrp = MXNU.CaricaStringaRes(24013) & " " & vntCodGrp
            'aggiungo nodo gruppo
            Set nodX = trwUtenti.Nodes.Add("gruppi", tvwChild, KeyGruppoGet(vntCodGrp), vntDscGrp, "gruppo")
            nodX.Tag = "G" & vntCodGrp
            Call nodX.EnsureVisible
            If (bolGruppoUtenti) Then
                'aggiungo nodi utenti
                Call TreeUtenti_Carica(trwUtenti, strSQL & vntCodGrp, nodX.key, vntCodGrp)
            End If
        End If
        bolEnd = Not MXDB.dbSuccessivo(HSS, TIPO_SNAPSHOT)
    Loop
    intq = MXDB.dbChiudiSS(HSS)
    'imposto utenti
    Set nodX = trwUtenti.Nodes.Add("Metodo98", tvwChild, "utenti", "Utenti", "entire")
    Call nodX.EnsureVisible
    
    If bolIncludiSupervisor Then
        strSQL = "SELECT UserID FROM TabUtenti ORDER BY UserID"
    Else
       strSQL = "SELECT UserID FROM TabUtenti WHERE Supervisor = 0 ORDER BY UserID"
    End If
    
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    bolEnd = MXDB.dbFineTab(HSS, TIPO_SNAPSHOT)
    Do While (Not bolEnd)
        vntCodUte = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "UserID", "")
        If (vntCodUte <> "") Then
            'aggiungo nodo utente
            
            ' Rif. anomalia n.ro 6183 (utente con codice numerico manda in errore l'add (esempio 1, 100, 231, ecc.))
            ' Stabilita la convenzione che in caso di codice numerico mettiamo un "underscore" come prefisso della stringa (es. 1 --> _1)
            If IsNumeric(vntCodUte) Then
                Set nodX = trwUtenti.Nodes.Add("utenti", tvwChild, "_" & vntCodUte, vntCodUte, "utente")
            Else
                Set nodX = trwUtenti.Nodes.Add("utenti", tvwChild, vntCodUte, vntCodUte, "utente")
            End If
            nodX.Tag = "U" & vntCodUte
            Call nodX.EnsureVisible
        End If
        bolEnd = Not MXDB.dbSuccessivo(HSS, TIPO_SNAPSHOT)
    Loop
    intq = MXDB.dbChiudiSS(HSS)
    
End Sub

Sub TreeUtenti_Carica(trwUtenti As TreeView, _
                        strSQL As String, _
                        strKeyParent As String, _
                        vntCodGrp As Variant)
Dim intq As Integer
Dim nodX As Node
Dim bolEnd As Boolean
Dim vntCodUte As Variant
Dim hSSUT As CRecordSet

    Set hSSUT = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    bolEnd = MXDB.dbFineTab(hSSUT, TIPO_SNAPSHOT)
    Do While (Not bolEnd)
        vntCodUte = MXDB.dbGetCampo(hSSUT, TIPO_SNAPSHOT, 0, "")
        If (vntCodUte <> "") Then
            'aggiungo nodo gruppo
            Set nodX = trwUtenti.Nodes.Add(strKeyParent, tvwChild, KeyUtenteGet(vntCodGrp, vntCodUte), vntCodUte, "utente")
            nodX.Tag = "U" & vntCodUte
        End If
        bolEnd = Not MXDB.dbSuccessivo(hSSUT, TIPO_SNAPSHOT)
    Loop
    intq = MXDB.dbChiudiSS(hSSUT)

End Sub

Function ValidaUtenteGruppo(vntValida As Variant, strKey As String) As Boolean
Dim intq As Integer
Dim strSQL As String
Dim HSS As CRecordSet
    
    ValidaUtenteGruppo = False
    strKey = ""
    If (vntValida <> "") Then
        'cerco fra gli utenti
        strSQL = "SELECT UserID FROM TabUtenti WHERE UserID='" & vntValida & "'"
        Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        If (Not MXDB.dbFineTab(HSS, TIPO_SNAPSHOT)) Then
            strKey = "U" & MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "UserID", "")
            ValidaUtenteGruppo = True
        Else
            'cerco fra i gruppi
            strSQL = "SELECT Codice FROM TabGruppiUtente WHERE Codice=" & Val(vntValida)
            Set HSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
            If (Not MXDB.dbFineTab(HSS, TIPO_SNAPSHOT)) Then
                strKey = "G" & MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "Codice", 0)
                ValidaUtenteGruppo = True
            End If
            intq = MXDB.dbChiudiSS(HSS)
        End If
        intq = MXDB.dbChiudiSS(HSS)
    End If
    
End Function



Public Function ComponiChiave(ByVal Nome As Variant, ByVal Indice As Variant) As Variant
    If Indice <> "" Then
        ComponiChiave = Nome & SEP_INDICE & Indice
    Else
        ComponiChiave = Nome
    End If
End Function

Public Function ScomponiChiave(ByVal Chiave As Variant, Indice As Variant) As Variant
    Dim Pos As Integer
    Dim Nome As Variant
    
    Nome = Chiave
    Indice = ""
    Pos = InStr(1, Chiave, SEP_INDICE)
    Do While Pos > 0
        Nome = Left(Chiave, Pos - 1)
        Indice = Mid(Chiave, Pos + 1)
        Pos = InStr(Pos + 1, Chiave, SEP_INDICE)
    Loop
    If IsNumeric(Indice) Then
        ScomponiChiave = Nome
    Else
        ScomponiChiave = Chiave
        Indice = ""
    End If
End Function

Public Function CheckControlEnabled(ctrls As Object) As Boolean
    Dim c As Control
    
    For Each c In ctrls
        If Not (TypeName(c) = "MWEtichetta" Or TypeName(c) = "MWLinguetta" Or TypeName(c) = "Label" Or TypeName(c) = "MWSchedaBox" Or TypeName(c) = "Image" Or TypeName(c) = "Frame" Or TypeName(c) = "cpvPicScroll" Or TypeName(c) = "ImageList" Or TypeName(c) = "XPToolButton") Then
            If c.Enabled And c.TabStop Then
                CheckControlEnabled = True
                Exit For
            End If
        End If
    Next
End Function

'Restituisce true se su mw.ini è stato attivato lo schedulare con Synapse (con schSynapse=1), false altrimenti
Function IsSchedSynapseActive() As Boolean
    Dim i As Integer
    
    'Prova a leggere dall'impostazione utente oppure globale
    i = Cast(MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", MXNU.UtenteSistema, "schSynapse", "-1"), vbInteger)
    If i = -1 Then
        i = Cast(MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "METODOW", "schSynapse", "-1"), vbInteger)
    End If
    
    'Se non è stato impostato lo mette per default a disattivato
    If i > 0 Then
        IsSchedSynapseActive = True
    Else
        IsSchedSynapseActive = False
    End If
End Function

