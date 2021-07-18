VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{B05A5950-FEDD-4618-B00B-87FEE9CE9470}#1.0#0"; "MXKIT.OCX"
Object = "{CFB3BBD6-56FA-41C3-A707-7CBC9EE47A51}#1.0#0"; "MXCTRL.OCX"
Begin VB.Form frmFiltroTabella 
   Appearance      =   0  'Flat
   Caption         =   "filtro tabella - standard -"
   ClientHeight    =   6750
   ClientLeft      =   720
   ClientTop       =   1065
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   10600
   Icon            =   "frmFiltroTabella.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6750
   ScaleWidth      =   11070
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      WhatsThisHelpID =   21055
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "filtro"
   End
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   6
      Top             =   0
      WhatsThisHelpID =   21120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Dati Selez."
   End
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      WhatsThisHelpID =   21108
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Totali"
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11245
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      ScaleWidth      =   11055
      ScaleHeight     =   6375
      Begin MXKit.ctlImpostazioni ctlImpFiltro 
         Height          =   555
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   60
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   979
      End
      Begin FPSpreadADO.fpSpread ssFiltroDati 
         Height          =   5415
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   10695
         _Version        =   196608
         _ExtentX        =   18865
         _ExtentY        =   9551
         _StockProps     =   64
         EditEnterAction =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         NoBeep          =   -1  'True
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "frmFiltroTabella.frx":0442
         UnitType        =   2
         VisibleCols     =   4
      End
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6360
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11218
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      ScaleWidth      =   11055
      ScaleHeight     =   6360
      Begin FPSpreadADO.fpSpread ssTotali 
         Height          =   6015
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   10695
         _Version        =   196608
         _ExtentX        =   18865
         _ExtentY        =   10610
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   3
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   30
         MaxRows         =   30
         NoBeep          =   -1  'True
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmFiltroTabella.frx":1C1A
         UnitType        =   2
         VisibleCols     =   30
         VisibleRows     =   30
         VScrollSpecial  =   -1  'True
      End
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6360
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11218
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      ScaleWidth      =   11055
      ScaleHeight     =   6360
      Begin FPSpreadADO.fpSpread ssRisult 
         Height          =   6015
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   10695
         _Version        =   196608
         _ExtentX        =   18865
         _ExtentY        =   10610
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   3
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   30
         MaxRows         =   999999
         NoBeep          =   -1  'True
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmFiltroTabella.frx":1E56
         UnitType        =   2
         VisibleCols     =   30
         VisibleRows     =   30
         VScrollSpecial  =   -1  'True
      End
   End
End
Attribute VB_Name = "frmFiltroTabella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine

'======================
'       costanti
'======================
Const SchFiltro = 0
Const schTabella = 1
Const schTotali = 2

'======================
'   classi
'======================
Private WithEvents xTabella As MXKit.CTabelle
Attribute xTabella.VB_VarHelpID = -1
Public WithEvents xFiltro As MXKit.CFiltro
Attribute xFiltro.VB_VarHelpID = -1
Public WithEvents MWAgt1 As MXKit.CAgenteAuto
Attribute MWAgt1.VB_VarHelpID = -1
Private AgenteTab(1) As String
Public FormProp As New CFormProp
Public MlngHlpTabella As Long

'======================
'       variabili
'======================
'variabili di stato della form
Dim mIntSchOnTop As Integer
Dim MlngButtonMask As Long
Dim mVetAgtTab(0 To 1) As String
'variabili per visione
Dim mStrNomeFiltro As String 'nome filtro
Dim mStrNomeTabella As String ' nome def tabella
Dim mBolInizializza As Boolean  'risultato dell'inizializzazione
Dim mBolFiltro As Boolean  'la visione ha filtro si/no
Dim mStrLogFile As String

'variabile da impostare per avere il controllo del valore nullo sulla colonna
Public IntCheckCol As Long
Public bolFoglioFisso As Boolean

Private strValCellaPrec As String
Private objDaEseguire As Object



'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
'           FUNZIONI PUBBLICHE DELLA FORM
'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
'classe per l'inizializzazione
'Nome del filtro
'Nome del def per la tabella
'Codice risorsa ling. per la caprio della form..
Public Function Imposta(objSetTabella As Object, _
                        strNomeFiltro As Variant, _
                        strNomeDefTabella As Variant, _
                        intTitolo As Long) As Boolean

    Dim hWndVis As Long

    ' associo la classe da eseguire alla fine della selezione..
    mStrNomeFiltro = strNomeFiltro
    mStrNomeTabella = strNomeDefTabella
    'Imposto la classe da eseguire per le varie funzioni..
    Set objDaEseguire = objSetTabella

    Me.Caption = MXNU.CaricaStringaRes(intTitolo)
    Set xTabella = MXCT.CreaCTabelle
    Set xFiltro = MXFT.CreaCFiltro()
    'imposto la linguetta on top..
    mIntSchOnTop = schTabella
    Call Ling_GotFocus(SchFiltro)
    'inizializzo il filtro dei dati,..
    ssFiltroDati.Visible = xFiltro.InizializzaFiltro(mStrNomeFiltro, ssFiltroDati)
    ctlImpFiltro(0).Visible = ctlImpFiltro(0).Inizializza(MXDB, MXNU, MXVI, mStrNomeFiltro, xFiltro, Nothing, ssFiltroDati, hndDBArchivi, GIMP_FILTRONORMALE)
    ssFiltroDati.Enabled = True
    'Inizializzo la tabella
    xTabella.NOMETABELLA = strNomeDefTabella
    Call objDaEseguire.InitTabella(xTabella, 0, ssRisult)
    If xTabella.Inizializza(ssRisult, MWAgt1) Then
        xTabella.StrWheAgg = "1=0"
        Call xTabella.TApriTabella(True, False)
        Call xTabella.TChiudiTabella
    End If
    Call objDaEseguire.InitTabella(xTabella, 0, ssRisult)

    Ling(schTotali).Visible = objSetTabella.FlagEnableTotali
    Ling(schTotali).Enabled = False

    'mostro la finestra
    Call CentraFinestra(Me.hwnd)
'Inzializzazione Form per Metodo Evolus
Call CambiaColoriControlli(Me)   'Da Mettere PRIMA di mResize
On Local Error Resume Next
Set mResize = New MxResizer.ResizerEngine
If (Not mResize Is Nothing) Then
	Call mResize.Initialize(Me, , , , , 0, True, metodo, MXNU)
End If
Call CentraFinestra(Me.hWnd)
Call CambiaCharSet(Me)
On Local Error GoTo 0
    Call Me.Show
    On Local Error GoTo 0
    Exit Function
End Function

Private Sub Ling_GotFocus(Index As Integer)
    Dim intC As Integer
    If Index <> mIntSchOnTop Then
        Select Case Index
            Case schTabella
                Call ssSpreadClear(ssRisult)
                xTabella.StrWheAgg = xFiltro.SQLFiltro
                If bolFoglioFisso Then ssRisult.MaxRows = 999999
                Call xTabella.TApriTabella(False, False)
                If bolFoglioFisso Then ssRisult.MaxRows = ssRisult.DataRowCnt
                ssRisult.Row = 1
                ssRisult.Col = IntCheckCol
                strValCellaPrec = ssCellGetValue(ssRisult, ssRisult.Col, ssRisult.Row)
                ssRisult.NoBeep = True
                MlngButtonMask = BTN_REG_MASK + BTN_ANN_MASK + BTN_PREC_MASK + BTN_PRIMO_MASK + BTN_SUCC_MASK + BTN_ULTIMO_MASK
                Call MXNU.Attiva_Toolbar(hwnd, MlngButtonMask)
                xTabella.TabModificata = False
                Ling(schTotali).Enabled = objDaEseguire.FlagEnableTotali
            Case SchFiltro
                Ling(schTotali).Enabled = False
                Call CheckTabella
                If xTabella.TTabAperta Then Call xTabella.TChiudiTabella
                MlngButtonMask = BTN_STP_MASK
                Call MXNU.Attiva_Toolbar(hwnd, MlngButtonMask)
            Case schTotali
                If Not (Ling(schTotali).Enabled) Then Exit Sub
        End Select
        Scheda(Index).Visible = True
        Scheda(mIntSchOnTop).Visible = False
        Call Scheda(Index).ZOrder(vbBringToFront)
        Call Ling(Index).ZOrder(vbBringToFront)
        Call Ling(mIntSchOnTop).ZOrder(vbBringToFront)
        mIntSchOnTop = Index
        Me.KeyPreview = (mIntSchOnTop = schTabella)
    End If
    'Le porto idietro
    For intC = 0 To Ling.Count - 1
        Call Ling(intC).ZOrder(1)
    Next intC
    Call Ling(Index).ZOrder(1)
    Ling(Index).OnTop = True
'Per Metodo Evolus
Call CambiaZOrderLinguette(Me)
End Sub


Private Sub Form_Activate()
    Call MXNU.Attiva_Toolbar(hwnd, MlngButtonMask)
    Call MXNU.ImpostaFormAttiva(Me)
End Sub

Private Sub Form_Load()
    
    Me.HelpContextID = FormProp.FormID
    If Me.HelpContextID = 0 Then
        Me.HelpContextID = Me.MlngHlpTabella
    End If
    MlngButtonMask = BTN_STP_MASK
    Set MWAgt1 = MXAA.CreaCAgenteAuto()
    Call MXNU.LeggiRisorseControlli(Me)
    Call MXAA.RegistraEventiFrm(Me, MWAgt1)
    Call MWAgt1.RegistraAgenteFrm(Me)
    Set objDaEseguire = Nothing
    'imposto la colonna di check ad un numero che mai avro' impostato
    IntCheckCol = 99999
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call xFiltro.ImpostaSTSColSS(STSSALVA)
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

Private Sub CheckTabella()
    If xTabella.TabModificata Then
        If MXNU.MsgBoxEX(2260, vbCritical + vbYesNo, 1007) = vbYes Then
            Call xTabella.TsalvaTab(AgenteTab, MWAgt1)
        Else
            xTabella.TabModificata = False
        End If
    End If
    If xTabella.TTabAperta Then Call xTabella.TChiudiTabella
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If xTabella.TabModificata Then
        If MXNU.MsgBoxEX(2260, vbCritical + vbYesNo, 1007) = vbYes Then
            Call xTabella.TsalvaTab(AgenteTab, MWAgt1)
        Else
            xTabella.TabModificata = False
        End If
    End If
    Call xTabella.TSalvaImpostazioni
    Call xTabella.TChiudiTabella(True)
    Set xTabella = Nothing
    
    Set FormProp = Nothing
    Set MWAgt1 = Nothing
    
    Set objDaEseguire = Nothing
    
    If ctlImpFiltro(0).Visible Then Call ctlImpFiltro(0).Termina
    Call xFiltro.ImpostaSTSColSS(stsSalva)
    Set xFiltro = Nothing
    
    Set frmFiltroTabella = Nothing
End Sub

Public Function AzioniMetodo(setAzione As enmFunzioniMetodo, Optional varparametro As Variant) As Variant

    Select Case setAzione
        Case MetFInserisci:
        Case MetFRegistra:
            Call xTabella.TsalvaTab(AgenteTab, MWAgt1)
            If bolFoglioFisso Then
                ssRisult.MaxRows = ssRisult.DataRowCnt
            End If
        Case MetFAnnulla
            Call xTabella.TDelRec
        Case MetFPrecedente
            If xTabella.FoglioGotFocus Then SendKeys "{UP}"
        Case MetFSuccessivo
            If xTabella.FoglioGotFocus Then SendKeys "{DOWN}"
        Case MetFPrimo
            If xTabella.FoglioGotFocus Then SendKeys "^{HOME}"
        Case MetFUltimo
            If xTabella.FoglioGotFocus Then SendKeys "^{END}"
        Case MetFDettagli:
        Case MetFStampa:
            'MXSpread.ssSpreadStampa (ssFiltroDati,)
        Case MetFVisUtenteModifica:
            'Call MXCT.VisDatiUtenteModifica(me, mStrNomeTabella, "", "", "", SSExtra)
        Case MetFDettVisione
        Case MetFMostraCampiDBAnagr
        Case MetFVisDipendenze
        Case Else
    End Select

End Function

Private Sub Scheda_Paint(Index As Integer)
    Call SchedaOmbreggiaControlli(Scheda(Index))
End Sub

Private Sub ssRisult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'se la colonna e' quella impostata per il controllo, allora mi arrabbio se hai inserito 0 e rimetto il valore precedente..
    If Col = IntCheckCol Then
        If ssCellGetValue(ssRisult, Col, Row) = 0 Then
            Beep
            Call ssCellSetValue(ssRisult, Col, Row, strValCellaPrec)
        End If
        strValCellaPrec = ssCellGetValue(ssRisult, NewCol, NewRow)
    End If
End Sub

Private Sub xFiltro_ValidazionePers(ByVal strNomeValid As String, ByVal strNomeCmpValid As String, bolEseguiValStd As Boolean, vntNewValore As Variant)
    Select Case strNomeValid
        Case "VALID_ARTICOLO"
            Dim xCodArt As MXBusiness.CVArt
            Set xCodArt = MXART.CreaCVArt()
            xCodArt.Codice = vntNewValore
            If xCodArt.Valida(CHIEDIVAR_TUTTE, False, , 0, False) Then
                vntNewValore = xCodArt.Codice
            End If
            bolEseguiValStd = True
            Call xCodArt.Termina
            Set xCodArt = Nothing
    End Select
    Set xCodArt = Nothing
End Sub







'Per Metodo Evolus
Private Sub mResize_AfterResize()
    Call AvvicinaLing(Me)
End Sub

