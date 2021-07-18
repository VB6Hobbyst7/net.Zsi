VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{1DCAEDEB-A59C-4B86-9BEC-40D93F9BE1EC}#1.0#0"; "MXKit.ocx"
Object = "{7C8739CE-EBF0-4404-AB39-B44A00355F1A}#1.0#0"; "mxctrl.ocx"
Begin VB.Form FrmFiltro 
   Caption         =   "Filtro di Stampa"
   ClientHeight    =   6315
   ClientLeft      =   2070
   ClientTop       =   2445
   ClientWidth     =   11055
   Icon            =   "Filtro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   11055
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6315
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11139
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bevel           =   1
      LightColor      =   6908265
      BevelWidth      =   2
      ScaleWidth      =   11055
      ScaleHeight     =   6315
      Begin VB.CommandButton CmdI 
         Caption         =   "&i"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10650
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   780
         Width           =   255
      End
      Begin VB.CommandButton ComRicostruisci 
         Caption         =   "&R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10380
         TabIndex        =   4
         Top             =   780
         WhatsThisHelpID =   10340
         Width           =   255
      End
      Begin VB.CommandButton ComProcedi 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   5760
         WhatsThisHelpID =   25005
         Width           =   2235
      End
      Begin FPSpreadADO.fpSpread ssFiltro 
         Height          =   4335
         Left            =   360
         TabIndex        =   1
         Top             =   1260
         Width           =   10515
         _Version        =   524288
         _ExtentX        =   18547
         _ExtentY        =   7646
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DAutoSizeCols   =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   14
         MaxRows         =   20
         NoBeep          =   -1  'True
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarShowMax=   -1  'True
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   4194368
         SpreadDesigner  =   "Filtro.frx":0E42
         VisibleCols     =   6
         VisibleRows     =   16
         AppearanceStyle =   0
      End
      Begin VB.ComboBox cmb 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   2820
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   780
         Width           =   7515
      End
      Begin MXKit.ctlImpostazioni ctlImp 
         Height          =   510
         Left            =   1560
         TabIndex        =   5
         Top             =   120
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   900
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   1
         Left            =   480
         Top             =   780
         WhatsThisHelpID =   10074
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LightColor      =   -2147483624
         ShadowColor     =   -2147483624
         VAlign          =   1
         Caption         =   "Stp"
      End
   End
End
Attribute VB_Name = "FrmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1

Public WithEvents MWAgt1 As MXKit.CAgenteAuto
Attribute MWAgt1.VB_VarHelpID = -1
Public FormProp As New CFormProp 'Aggiunta in data 28/3/2000 Daniel
Attribute FormProp.VB_VarHelpID = -1

'Public MXFT As New MXKit.CAmbiente
Dim WithEvents objFiltro As MXKit.CFiltro
Attribute objFiltro.VB_VarHelpID = -1

Public strNomeFiltro As String
Public bolDaAgente As Boolean
Public SQLFiltro As String
Public Formule As MXKit.ParAgg
Public DSNDitta As String
Public FiltroConTemp As Boolean

'Dim objCRW As MXKit.CCrw

Dim Maschera&
Dim sngLarghSSDesign As Single
Dim McolObjCrw As New Collection

Dim Nomi_Stampe(0 To 99) As String
Dim Nomi_SubReport(0 To 99) As String
Dim ImpComune As Boolean   'Indica se l'impostazione caricata è comune per tutti gli utenti (Usato per l'eventuale eliminazione)
Dim MbolInElaborazione As Boolean

Public Property Get xFiltro() As MXKit.CFiltro

    Set xFiltro = objFiltro
    
End Property

Private Sub DefLingua()
    Call MXNU.LeggiRisorseControlli(Me)
    ComRicostruisci.Caption = "&R"
'    Etc(0).Caption = MXNU.CaricaStringaRes(10073)
'    Etc(1).Caption = MXNU.CaricaStringaRes(10074)
'    ComImpost(0).Caption = MXNU.CaricaStringaRes(25003)
'    ComImpost(1).Caption = MXNU.CaricaStringaRes(25004)
'    ComProcedi.Caption = MXNU.CaricaStringaRes(25005)
'    ComAnnulla.Caption = MXNU.CaricaStringaRes(25006)
End Sub

Public Function AzioniMetodo(setAzione As enmFunzioniMetodo, Optional varparametro As Variant) As Variant
    
    Select Case setAzione
        Case MetFSchedulaOperazione
            If MAccessi.IsSchedSynapseActive Then
                Call NuovaSchedulazioneSyn
            Else
                Call NuovaSchedulazione
            End If
            
        Case Else
    End Select

End Function

Private Function GetParameterXml(reportname As String, _
        filtroSQL As String, periferica As String, _
        nomestampante As String, recipients As String, _
        oggetto As String, body As String, filepath As String, _
        formatofile As String) As String
        
Dim oXml As MSXML2.DOMDocument
Dim oPiNode As MSXML2.IXMLDOMProcessingInstruction
Dim oRootNode As MSXML2.IXMLDOMNode
Dim oParamNode As MSXML2.IXMLDOMNode

    Set oXml = New MSXML2.DOMDocument
    
    'processing istruction
    Set oPiNode = oXml.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    oXml.appendChild oPiNode
    'nodo root
    Set oRootNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameters", vbNullString)
    oXml.appendChild oRootNode
    'nodi parametri
    'parametro report
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = reportname
    Call XmlCreateAttribute(oParamNode, "name", "Report")
    oRootNode.appendChild oParamNode

    'parametro nome stampante
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = nomestampante
    Call XmlCreateAttribute(oParamNode, "name", "NomeStampante")
    oRootNode.appendChild oParamNode

    'parametro filtrosql
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = filtroSQL
    Call XmlCreateAttribute(oParamNode, "name", "FiltroSQL")
    oRootNode.appendChild oParamNode

     'parametro metodointerop
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = MXNU.DittaAttiva & ";" & MXNU.PasswordUtente & ";" & MXNU.UtenteAttivo & ";"
    Call XmlCreateAttribute(oParamNode, "name", "MetodoInterop")
    oRootNode.appendChild oParamNode
    
    'parametro periferica
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = periferica
    Call XmlCreateAttribute(oParamNode, "name", "Periferica")
    oRootNode.appendChild oParamNode
    
    'parametro fileoutput
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = filepath
    Call XmlCreateAttribute(oParamNode, "name", "FileOutput")
    oRootNode.appendChild oParamNode
    
    'parametro formatofile
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = formatofile
    Call XmlCreateAttribute(oParamNode, "name", "FormatoFile")
    oRootNode.appendChild oParamNode

    'parametro recipients
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = recipients
    Call XmlCreateAttribute(oParamNode, "name", "Recipients")
    oRootNode.appendChild oParamNode

    'parametro oggetto
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = oggetto
    Call XmlCreateAttribute(oParamNode, "name", "Oggetto")
    oRootNode.appendChild oParamNode

    'parametro body
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = body
    Call XmlCreateAttribute(oParamNode, "name", "Body")
    oRootNode.appendChild oParamNode
    
    'parametro XMLCFiltro
    Set oParamNode = oXml.createNode(MSXML2.NODE_ELEMENT, "parameter", vbNullString)
    oParamNode.text = SerializeCFiltroToXML(xFiltro)
    Call XmlCreateAttribute(oParamNode, "name", "XMLCFiltro")
    oRootNode.appendChild oParamNode

    Set oPiNode = Nothing
    Set oRootNode = Nothing
    Set oParamNode = Nothing
    
    GetParameterXml = oXml.xml
    Set oXml = Nothing
End Function

Private Function GetPeriferica4Syn(periferica As String) As Integer
    Select Case periferica
        Case "E-Mail"
            GetPeriferica4Syn = 1
        Case "File"
            GetPeriferica4Syn = 2
        Case "Stampante"
            GetPeriferica4Syn = 0
    End Select
End Function

Private Function GetFormatoString(Formato As Integer) As String

    Select Case Formato
        Case 31
            GetFormatoString = "pdf"
        Case 32
            GetFormatoString = "html"
        Case 8
            GetFormatoString = "text"
        Case 37
            GetFormatoString = "xml"
        Case 14
            GetFormatoString = "ms word"
        Case 36
            GetFormatoString = "ms excel"
        Case Else
            GetFormatoString = "pdf"
    End Select
    
End Function

Private Function NuovaSchedulazioneSyn() As Boolean
    Dim executor As ISmartTagExecutor
    Dim objFiltroStp As MXKit.CFiltro
    Dim filtroSQL As String
    Dim periferica As String
    Dim reportname As String
    Dim body As String
    Dim recipients As String
    Dim oggetto As String
    Dim filepathoutput As String
    Dim formatofile As String
    Dim printername As String
    Dim objCRW As MXKit.CCrw
    
    Set executor = CreateObject("MxSynapseBridge.SynapseExecutor")
    
     'inizializzo oggetto CCrw
    Set objCRW = MXCREP.CreaCCrw()
    objCRW.ClearOpzioniStp
    objCRW.Titolo = cmb(1).text
    objCRW.Filerpt = Nomi_Stampe(cmb(1).ItemData(cmb(1).listIndex))
    objCRW.FiltroEMailAnagr = (objFiltro.FiltroFax <> "")
    objCRW.OpzioniForm = STP_TUTTE - STP_ANTEPRIMA 'blocca il pulsante anteprima!
    objCRW.MostraFrmStampa

    If objCRW.periferica <> "Video" Then
        'q uesto punto posso lanciare
        Set objFiltroStp = objFiltro
        filtroSQL = objFiltroStp.SQLFiltro
        periferica = GetPeriferica4Syn(objCRW.periferica)
        reportname = Replace(Nomi_Stampe(cmb(1).ItemData(cmb(1).listIndex)), "%PATHPGM%", MXNU.PercorsoStampe)
        printername = objCRW.Stampante.nomestampante
        If objCRW.periferica = "File" Then
            'memorizzare formato
            formatofile = GetFormatoString(objCRW.FormatoStampaSuFile)
            'memorizzare nome file
            filepathoutput = objCRW.FileOutput
        ElseIf objCRW.periferica = "E-Mail" Then
            recipients = objCRW.DestinatariEMail
            'memorizzare oggetto msg
            oggetto = objCRW.OggettoEMail
            'memorizzare testo msg
            body = objCRW.TestoEMail
        End If
    End If
    
    Dim parameters As String
    Dim odom As New MSXML2.DOMDocument
    odom.Load (MXNU.PercorsoPgm & "\MxSchedulerSynapse.config")
    
    Dim onodetask As MSXML2.IXMLDOMNode
    Set onodetask = odom.selectSingleNode("configuration/operations/operation[@id='2']")
    If Not onodetask Is Nothing Then
        'ritorno la definizione
        Dim onodedef As MSXML2.IXMLDOMNode
        Set onodedef = onodetask.selectSingleNode("definition")
        Dim def As String
        def = onodedef.xml
        
        If executor.Execute(mMetodoInterop, def, GetParameterXml(reportname, _
                        filtroSQL, periferica, printername, recipients, _
                        oggetto, body, filepathoutput, formatofile)) Then

        End If
    End If
End Function

' modifica del 25/02/2002 - Utilizzato il nuovo job scheduler
Private Function NuovaSchedulazione() As Boolean

#If BATCH <> 1 And ISM98SERVER <> 1 And TOOLS <> 1 Then
    Dim bolRes As Boolean
    Dim objSchedula As MxScheduler.clsScheduler
    Dim objOperDb As clsOperDb
    Dim objCRW As MXKit.CCrw
    Dim strImpStp As String

    bolRes = False
    On Local Error GoTo ERR_NuovaSchedulazione
    
    ' impostazione parametri operazione schedulata
    Set objSchedula = New MxScheduler.clsScheduler
    If objSchedula.Inizializza(MXNU, Command()) Then
        Set objOperDb = objSchedula.CreaOperazione("ESEGUISTAMPA", True)
        objOperDb.Descrizione = ctlImp.NomeImpostazione
        Call objOperDb.SetRiga("FILTRO", 1, strNomeFiltro)
        Call objOperDb.SetRiga("FILTRO", 3, Nomi_Stampe(cmb(1).ItemData(cmb(1).listIndex)))
        Call objOperDb.SetRiga("FILTRO", 5, ctlImp.NomeImpostazione)
    
        'inizializzo oggetto CCrw
        Set objCRW = MXCREP.CreaCCrw()
        objCRW.ClearOpzioniStp
        objCRW.Titolo = cmb(1).text
        objCRW.Filerpt = Nomi_Stampe(cmb(1).ItemData(cmb(1).listIndex))
        objCRW.FiltroEMailAnagr = (objFiltro.FiltroFax <> "")
        objCRW.OpzioniForm = STP_TUTTE - STP_ANTEPRIMA 'blocca il pulsante anteprima!
        objCRW.MostraFrmStampa
        
        'memorizzare tipo periferica
        If objCRW.periferica <> "Video" Then
            Call objOperDb.SetRiga("OPZIONI", 1, objCRW.periferica)
        
            With objCRW.Stampante
                strImpStp = .nomestampante & "^" & .NomeDriver & "^" & .NomePorta & "^" _
                    & .dmOrientation & "^" & .dmPaperSize & "^" & .dmPrintQuality & "^" _
                    & .flags & "^" & .dmCollate & "^" & .dmColor & "^" & .dmDuplex & "^" _
                    & .dmFormName
            End With
            'strImpStp = ImpostaStampante
            Call objOperDb.SetRiga("STAMPANTE", 1, strImpStp)
            Call objOperDb.SetRiga("STAMPANTE", 2, Nomi_SubReport(cmb(1).ItemData(cmb(1).listIndex)))
            
            If objCRW.periferica = "ZFax" Then
                If objCRW.FiltroEMailAnagr Then
                    'memorizzare ruolo (altrimenti non serve memoriz questa informazione)
                    Call objOperDb.SetRiga("OPZIONI", 2, objCRW.Ruoli.Ruolo)
                End If
            
            ElseIf objCRW.periferica = "File" Then
                'memorizzare formato
                Call objOperDb.SetRiga("FILE", 1, objCRW.FormatoStampaSuFile)
                'memorizzare nome file
                Call objOperDb.SetRiga("FILE", 2, objCRW.FileOutput)
            
            ElseIf objCRW.periferica = "E-Mail" Then
                If objCRW.FiltroEMailAnagr Then
                    'memorizzare ruolo (dal ruolo si becca l'indirizzo email)
                    Call objOperDb.SetRiga("OPZIONI", 2, objCRW.Ruoli.Ruolo)
                Else
                    'memorizzare solo destinatari email
                    Call objOperDb.SetRiga("EMAIL", 3, objCRW.DestinatariEMail)
                End If
                'memorizzare oggetto msg
                Call objOperDb.SetRiga("EMAIL", 1, objCRW.OggettoEMail)
                'memorizzare testo msg
                Call objOperDb.SetRiga("EMAIL", 2, objCRW.TestoEMail)
            End If
        
            Call objSchedula.NewOperation(objOperDb, False)
            bolRes = True
        End If
        
        Set objCRW = Nothing
    End If
    
END_NuovaSchedulazione:
    On Local Error GoTo 0
    Set objOperDb = Nothing
    Set objSchedula = Nothing
    NuovaSchedulazione = bolRes
    Exit Function
    
ERR_NuovaSchedulazione:
    bolRes = False
    Call MXNU.MsgBoxEX("NewSchedulazione" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, App.Title)
    Resume END_NuovaSchedulazione
    
#End If
    
End Function

Private Function ImpostaStampante() As String

Dim strResult As String
Dim objStp As MXKit.CStampante

    strResult = ""
    On Local Error GoTo ERR_ImpostaStampante
    Set objStp = New MXKit.CStampante
    With objStp
        .flags = PD_NONETWORKBUTTON + PD_RETURNDC + PD_HIDEPRINTTOFILE + PD_NOSELECTION + PD_PRINTSETUP
        .dmFields = DM_COPIES + DM_DUPLEX + DM_ORIENTATION + DM_PAPERSIZE + DM_PRINTQUALITY
        .nMaxPage = 1
        .nMinPage = 1
        If .Imposta(Me.hwnd) Then
            strResult = .nomestampante & "^" & .NomeDriver & "^" & .NomePorta & "^" & .dmOrientation & "^" & .dmPaperSize & "^" & .dmPrintQuality & "^" & .flags & "^" & .dmCollate & "^" & .dmColor & "^" & .dmDuplex & "^" & .dmFormName
        End If
    End With
                
END_ImpostaStampante:
    On Local Error GoTo 0
    Set objStp = Nothing
    ImpostaStampante = strResult
    Exit Function
                
ERR_ImpostaStampante:
    strResult = ""
    Call MXNU.MsgBoxEX("Function ImpostaStampante: " & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, App.Title)
    Resume END_ImpostaStampante

End Function


Private Sub cmb_Click(Index As Integer)
    Dim strFileHelp As String
    
    If Index = 1 Then
        'Call objFiltro.AttivaRigheFiltroAgg(cmb(Index).ListIndex)
        Call objFiltro.AttivaRigheFiltroAgg(Nomi_Stampe(cmb(Index).ItemData(cmb(Index).listIndex)))
        Call objFiltro.AttivaRigheSubReport(Nomi_SubReport(cmb(Index).ItemData(cmb(Index).listIndex)))
        ssFiltro.Visible = False
        'ssFiltro.Col = 1
        'ssFiltro.Row = ssFiltro.ActiveRow
        ssFiltro.Row = 1
        ssFiltro.Col = 1
        DoEvents
        ssFiltro.Action = SS_ACTION_GOTO_CELL

'        ssFiltro.AutoSize = True
'        ssFiltro.AutoSize = False

        ssFiltro.ReDraw = False
        DoEvents
        If ssFiltro.width > sngLarghSSDesign Then
            ssFiltro.width = sngLarghSSDesign
        End If
        ssFiltro.ReDraw = True
        
        'RIF. AN. #10449 rz
        If Not mResize Is Nothing Then
            Call mResize.Frm_Resize(True)
        End If

        'Me.width = Me.width + 1
        'DoEvents
        'Me.width = Me.width - 1


'        ssFiltro.AutoSize = True
'        ssFiltro.AutoSize = False
'        DoEvents
'        If ssFiltro.Width > sngLarghSSDesign Then
'            ssFiltro.Width = sngLarghSSDesign
'        End If
        ssFiltro.Visible = True
        'Me.Refresh
        Scheda_Paint 0
    
        strFileHelp = Mid$(Nomi_Stampe(cmb(1).listIndex), InStrRev(Nomi_Stampe(cmb(1).listIndex), "\") + 1)
        strFileHelp = Left$(strFileHelp, InStr(strFileHelp, ".")) & "htm"
    End If
End Sub


Private Sub CmdI_Click()
    
    Call frmInfoStp.MostraImpstazioniStampa(cmb(1), Nomi_Stampe(), objFiltro, strNomeFiltro)
    
End Sub

Private Sub ComProcedi_Click()
    If MbolInElaborazione Then Exit Sub   'Rif. Anomalia Nr. 7465
    
    ' Rif. anomalia #4414
    ' Se si decommenta questa riga per i click multipli la scheda anomalia agenti 4414 (esecuzione agente con ALT + Lettera sottolineata) NON si può risolvere.
    'ComProcedi.Enabled = False
    DoEvents
    MbolInElaborazione = True
    Dim q%, i%, objCRW As MXKit.CCrw
    If objFiltro.EseguiStampa Then
        Dim bolUsaTemp As Boolean
        Dim objFiltroStp As MXKit.CFiltro
        If cmb(1).listIndex < 0 Then
            MXNU.MsgBoxEX 1111, vbExclamation, 1007
            Exit Sub
        End If
        
        Set objCRW = MXCREP.CreaCCrw()
        If DSNDitta <> "" Then objCRW.DSNDitta = DSNDitta
        McolObjCrw.Add objCRW
        
        objCRW.ClearOpzioniStp
        objCRW.Titolo = cmb(1).text
        objCRW.Filerpt = Nomi_Stampe(cmb(1).ItemData(cmb(1).listIndex))
        Dim frmAnt As New FrmAnteprima
        DoEvents
        ComProcedi.Enabled = True
        
        '*******************************************************************************************************************
        'Gestione filtro che crea un temporaneo tramite una stored procedure lanciata da una riga di tipo Query del filtro
        '*******************************************************************************************************************
        If FiltroConTemp Then
            bolUsaTemp = (InStr(cmb(1).List(cmb(1).listIndex), "[T]") > 0)
            If bolUsaTemp Then
                If objFiltro.Query.Count > 0 Then
                    Call objFiltro.EseguiQuery(TQ_PRIMASTP)
                End If
            
                Set objFiltroStp = MXFT.CreaCFiltro()
                Call objFiltroStp.InizializzaFiltro
                Set objFiltroStp.ParAgg = objFiltro.ParAgg
                Set objFiltroStp.Raggruppa = objFiltro.Raggruppa
                Set objFiltroStp.Ordinamento = objFiltro.Ordinamento
            Else
                Set objFiltroStp = objFiltro
            End If
        Else
            Set objFiltroStp = objFiltro
        End If
        '*******************************************************************************************************************

        frmLog.ClearLog
        frmLog.UseExistingLogPanel = False
        objCRW.AttendiAnteprima = (bolDaAgente)   'Anomalia 11490
        q = objCRW.Stampa(objFiltroStp, frmAnt, True)
        Set frmAnt = Nothing '
        If objCRW.StpFiltro Then Call objFiltro.StampaFiltro(objCRW)
        If FiltroConTemp And bolUsaTemp Then   'Eseguo le eventuali query post-stampa (in quanto avendo utilizzato un filtro diverso senza la where, non le contiene)
            If objFiltro.Query.Count > 0 Then
                Call objFiltro.EseguiQuery(TQ_DOPOSTP)
            End If
        End If
        DoEvents
        Set objFiltroStp = Nothing
        For i = 1 To McolObjCrw.Count   'Scarico l'istanza di objCrw appena creata per effettuare la stampa
            If McolObjCrw(i) Is objCRW Then
                McolObjCrw.Remove i
                Exit For
            End If
        Next i
        Set objCRW = Nothing
        
    Else
        If objFiltro.CtrlCampiObbligatori Then
            q = objFiltro.EseguiQuery(0)
            If q And objFiltro.QueryDiAnnullamento Then
                SendKeys "^{F4}"
            End If
        End If
    End If
    If bolDaAgente Then
        Dim objFE As CFiltriElenco
        Dim strNomeFormula As Variant
        SQLFiltro = objFiltro.SQLFiltro
        Set Formule = objFiltro.ParAgg
        On Local Error Resume Next
        For Each objFE In objFiltro.FiltriElenco
            q = objFiltro.IdFiltro2Riga(objFE.IDRiga)
            If ssFiltro.GetText(COLFORMULA, q, strNomeFormula) Then
                strNomeFormula = Replace(strNomeFormula, "@", "")
                Call Formule.Add(CStr(strNomeFormula), CStr(strNomeFormula), objFE.SQLQuery, "", True, CStr(strNomeFormula))
            End If
        Next
        Set objFE = Nothing
        If Not (objFiltro.EseguiStampa) Then
            SendKeys "^{F4}"
        End If
    End If
    ' Rif. anomalia #4414
    ' Se si decommenta questa riga per i click multipli la scheda anomalia agenti 4414 (esecuzione agente con ALT + Lettera sottolineata) NON si può risolvere.
    'ComProcedi.Enabled = True
    MbolInElaborazione = False
End Sub

Private Sub ComRicostruisci_Click()
    Dim q%
    q = MXNU.MsgBoxEX(1114, vbYesNo + vbQuestion, 1007)
    If q = vbYes Then
        CaricaListaStampe objFiltro.NomeFiltro, cmb(1), Nomi_Stampe(), Nomi_SubReport(), True
        'If cmb(1).ListCount > 0 Then cmb(1).ListIndex = 0
        If MXNU.UsaEuro Then
            Call PosizioneStpEuro(cmb(1))
        Else
            If cmb(1).ListCount > 0 Then cmb(1).listIndex = 0   'Aggiunto IF: Rif. Anomalia Nr. 7349
        End If
    End If

End Sub

Private Sub ctlImp_DopoCaricamento()
    Dim vetID() As Long
    Dim i%
    Select Case UCase(strNomeFiltro)
        'Rif. Anomalia Nr. 3914
        'Caricamento di alcuni valori dalla query di default indipendentemente dalla impostazione caricata
        Case "STP_RIEPREGIVA"
            ReDim vetID(1 To 2) As Long
            vetID(1) = 6
            vetID(2) = 11
            For i = 1 To UBound(vetID)
                objFiltro.ImpostaDefaultCella COLVALOREDA, objFiltro.IdFiltro2Riga(vetID(i))
                objFiltro.ImpostaDefaultCella COLVALOREA, objFiltro.IdFiltro2Riga(vetID(i))
            Next i
            Erase vetID
        Case "STP_LQIVA"   'Rif. Anomalia 4970
            ReDim vetID(1 To 5) As Long
            vetID(1) = 33
            vetID(2) = 34
            vetID(3) = 6
            vetID(4) = 18
            vetID(5) = 10
            For i = 1 To UBound(vetID)
                objFiltro.ImpostaDefaultCella COLVALOREDA, objFiltro.IdFiltro2Riga(vetID(i))
                objFiltro.ImpostaDefaultCella COLVALOREA, objFiltro.IdFiltro2Riga(vetID(i))
            Next i
            Erase vetID
            
    End Select
    
End Sub

Private Sub Form_Activate()
    Call MXNU.Attiva_Toolbar(hwnd, Maschera)
    Call MXNU.ImpostaFormAttiva(Me)
End Sub

Private Sub Form_Load()
    Dim res%
    
    metodo.MousePointer = vbHourglass
    
    Me.HelpContextID = FormProp.FormID
    
    Set MWAgt1 = MXAA.CreaCAgenteAuto()
    res = MXAA.RegistraEventiFrm(Me, MWAgt1)

    DefLingua
    'If MXFT.inizializza(MXNU, MXDB, MXSE, metodo, hndDBArchivi, hndDBUtenti) Then
        sngLarghSSDesign = ssFiltro.width
        Set objFiltro = MXFT.CreaCFiltro()
        If objFiltro.InizializzaFiltro(strNomeFiltro, ssFiltro) Then

            'Set objCrw = MXCREP.CreaCCrw()
            'If DSNDitta <> "" Then objCrw.DSNDitta = DSNDitta
            If objFiltro.DescrizioneFiltro <> "" Then
                Me.Caption = objFiltro.DescrizioneFiltro
            Else
                Me.Caption = MXNU.CaricaStringaRes(23003)
            End If

            ssFiltro.VisibleCols = 6

            CaricaListaStampe objFiltro.NomeFiltro, cmb(1), Nomi_Stampe(), Nomi_SubReport(), False
            'If cmb(1).ListCount > 0 Then cmb(1).ListIndex = 0
            If MXNU.UsaEuro Then
                Call PosizioneStpEuro(cmb(1))
            Else
                If cmb(1).ListCount > 0 Then cmb(1).listIndex = 0   'Aggiunto IF: Rif. Anomalia Nr. 7349
            End If

            If Not ctlImp.Inizializza(MXDB, MXNU, MXVI, strNomeFiltro, objFiltro, cmb(1), ssFiltro, hndDBArchivi) Then
                ctlImp.Visible = False
            End If
            'strPercorsoStampaGrafica = MXNU.PercorsoStampe & "\laser\"
            'strPercorsoStampaTesto = MXNU.PercorsoStampe & "\testo\"
            'strPercorsoStampaModuli = MXNU.PercorsoStampe & "\Moduli\"
            'ssFiltro.AutoSize = True
            'ssFiltro.AutoSize = False
            If ssFiltro.width > sngLarghSSDesign Then ssFiltro.width = sngLarghSSDesign
            If Not objFiltro.EseguiStampa Then
                etc(1).Visible = False
                cmb(1).Visible = False
                ComRicostruisci.Visible = False
            End If
            CentraFinestra Me.hwnd
            
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
        Else
            MXNU.MsgBoxEX 1110, vbExclamation, 1007
            metodo.MousePointer = vbDefault
            Unload Me
            Exit Sub
        End If
    'End If
        
    CmdI.Visible = cmb(1).Visible
    
    metodo.MousePointer = vbDefault

End Sub

Public Sub MetInserisci()

End Sub

Public Sub MetDettagli()

End Sub

Public Sub MetRegistra()

End Sub

Public Sub MetAnnulla()

End Sub

Public Sub MetPrimo()

End Sub

Public Sub MetPrecedente()

End Sub

Public Sub MetSuccessivo()

End Sub

Public Sub MetUltimo()


End Sub

Public Sub MetStampa()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MbolInElaborazione Then
        Cancel = True
    Else
        If McolObjCrw.Count > 0 Then
            Dim objCRW As MXKit.CCrw
            For Each objCRW In McolObjCrw
                Cancel = objCRW.InStampa
                If Cancel Then Exit For
            Next objCRW
        End If
    End If
    'If Not (objCrw Is Nothing) Then
    '    Cancel = objCrw.InStampa
    'End If
    
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
    Call ctlImp.Termina
    Call objFiltro.ImpostaSTSColSS(MXKit.stsSalva)
    Set MWAgt1 = Nothing
    Set Formule = Nothing
    Set objFiltro = Nothing
    If McolObjCrw.Count > 0 Then
        Dim objCRW As MXKit.CCrw
        For Each objCRW In McolObjCrw
            Set objCRW = Nothing
        Next objCRW
    End If
    Set McolObjCrw = Nothing
    'Set objCrw = Nothing
    Set FormProp = Nothing 'Aggiunta in data 28/3/2000 Daniel
    Set FrmFiltro = Nothing
End Sub

Private Sub objFiltro_ValidazionePers(ByVal strNomeValid As String, ByVal strNomeCmpValid As String, bolEseguiValStd As Boolean, vntNewValore As Variant)
    Call ValidPersFiltri(strNomeValid, strNomeCmpValid, bolEseguiValStd, vntNewValore)
End Sub

Private Sub Scheda_Paint(Index As Integer)
   SchedaOmbreggiaControlli scheda(Index)
    'SCHEDADRAWDROPSHADOWS Scheda(Index).hwnd, ssFiltro, 2
End Sub











'Per Metodo Evolus
Private Sub mResize_AfterResize()
    Call AvvicinaLing(Me)
End Sub

'*************************************************************************
'TODO: Eliminare questo codice una volta uscita la verione 9.0
'*************************************************************************
Private Function SerializeCFiltroToXML(xFiltro As CFiltro) As String
    If xFiltro Is Nothing Then
        SerializeCFiltroToXML = ""
        Exit Function
    End If
    
    Dim doc As New DOMDocument
    Dim xe As IXMLDOMElement
    Dim xc As IXMLDOMElement
    Dim Col As IXMLDOMElement
    Dim o As IXMLDOMElement
    Dim i As Integer
    
    Set xe = doc.createElement("CFiltro")
    Call doc.appendChild(xe)
    
    Set xc = AddChildXMLElement(doc, xe, "CParAgg")
    i = 1
    Dim pa As CParAgg
    For Each pa In xFiltro.ParAgg
        Set o = AddChildXMLElement(doc, xc, "CParAgg" & i)
        Call AddChildXMLElement(doc, o, "Key", pa.key)
        Call AddChildXMLElement(doc, o, "NomeFormula", pa.key)
        Call AddChildXMLElement(doc, o, "ValoreFormula", pa.ValoreFormula)
        Call AddChildXMLElement(doc, o, "NomiSubReport", pa.NomiSubReport)
        Call AddChildXMLElement(doc, o, "FormulaCustom", Bool2XML(pa.FormulaCustom))
        Call AddChildXMLElement(doc, o, "sKey", pa.key)
    
        i = i + 1
    Next
    
    Set xc = AddChildXMLElement(doc, xe, "COrdinamento")
    i = 1
    Dim ord As COrdinamento
    For Each ord In xFiltro.Ordinamento
        Set o = AddChildXMLElement(doc, xc, "COrdinamento" & i)
        Call AddChildXMLElement(doc, o, "Key", ord.key)
        Call AddChildXMLElement(doc, o, "EseguiOrd", Bool2XML(ord.EseguiOrd))
        Call AddChildXMLElement(doc, o, "ListaCampi", ord.ListaCampi)
        Call AddChildXMLElement(doc, o, "NomiSubReport", ord.NomiSubReport)
        Call AddChildXMLElement(doc, o, "sKey", ord.key)
    
        i = i + 1
    Next
    
    Set xc = AddChildXMLElement(doc, xe, "CRaggruppa")
    i = 1
    Dim rag As CRaggruppa
    For Each rag In xFiltro.Raggruppa
        Set o = AddChildXMLElement(doc, xc, "CRaggruppa" & i)
        Call AddChildXMLElement(doc, o, "Key", rag.key)
        Call AddChildXMLElement(doc, o, "CampoCodice", rag.CampoCodice)
        Call AddChildXMLElement(doc, o, "CampoDescrizione", rag.CampoDescrizione)
        Call AddChildXMLElement(doc, o, "CampoDBGruppo", rag.CampoDBGruppo)
        Call AddChildXMLElement(doc, o, "EseguiRaggruppa", Bool2XML(rag.EseguiRaggruppa))
        Call AddChildXMLElement(doc, o, "FormulaCodice", rag.FormulaCodice)
        Call AddChildXMLElement(doc, o, "FormulaDescrizione", rag.FormulaDescrizione)
        Call AddChildXMLElement(doc, o, "NumeroGruppo", rag.NumeroGruppo)
        Call AddChildXMLElement(doc, o, "NomiSubReport", rag.NomiSubReport)
        Call AddChildXMLElement(doc, o, "sKey", rag.key)
        
        Dim j As Integer
        Dim f As Variant
        Set Col = AddChildXMLElement(doc, o, "colFormuleCustom")
        j = 1
        For Each f In rag.colFormuleCustom
            Call AddChildXMLElement(doc, Col, "Item" & j, CStr(f))
            
            j = j + 1
        Next
        
        Set Col = AddChildXMLElement(doc, o, "colValFormuleCustom")
        j = 1
        For Each f In rag.colValFormuleCustom
            Call AddChildXMLElement(doc, Col, "Item" & j, CStr(f))
            
            j = j + 1
        Next
    
        i = i + 1
    Next
    
    SerializeCFiltroToXML = doc.xml
End Function

Private Function AddChildXMLElement(doc As DOMDocument, Node As IXMLDOMElement, NAME As String, Optional value As String = "") As IXMLDOMElement
    Dim xe As IXMLDOMNode
    
    Set xe = doc.createElement(NAME)
    Call Node.appendChild(xe)
    
    If value <> "" Then
        xe.text = value
    End If
    
    Set AddChildXMLElement = xe
End Function

Private Function Bool2XML(value As Boolean) As String
    If value = True Then
        Bool2XML = "true"
    Else
        Bool2XML = "false"
    End If
End Function

'*************************************************************************
'TODO: FINE
'*************************************************************************

