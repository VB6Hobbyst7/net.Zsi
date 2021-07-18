VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{7BF04B61-D576-4084-9107-4EA0960A7556}#1.0#0"; "MXCtrl.ocx"
Object = "{ED9F8BCC-09C1-4AA6-8BA2-9956D7545976}#1.0#0"; "MXKit.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.UserControl GeneraCodici 
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13050
   ControlContainer=   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   13050
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      WhatsThisHelpID =   21028
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Filtro"
      First           =   -1  'True
   End
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      Index           =   1
      Left            =   1710
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      WhatsThisHelpID =   21028
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ordini"
      First           =   -1  'True
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6495
      Index           =   0
      Left            =   15
      TabIndex        =   4
      Top             =   315
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   11456
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LightColor      =   6908265
      BevelWidth      =   2
      ScaleWidth      =   12915
      ScaleHeight     =   6495
      Begin VB.CommandButton ComProcedi 
         Caption         =   "Procedi"
         Height          =   375
         Index           =   0
         Left            =   5408
         TabIndex        =   3
         Top             =   5970
         WhatsThisHelpID =   25005
         Width           =   2235
      End
      Begin MXKit.ctlImpostazioni ctlImp 
         Height          =   555
         Left            =   705
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   979
      End
      Begin FPSpreadADO.fpSpread ssFiltro 
         Height          =   5010
         Left            =   135
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   765
         Width           =   12285
         _Version        =   524288
         _ExtentX        =   21669
         _ExtentY        =   8837
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoBeep          =   -1  'True
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "GeneraCodici.ctx":0000
         AppearanceStyle =   0
      End
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6495
      Index           =   1
      Left            =   15
      TabIndex        =   5
      Top             =   315
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   11456
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LightColor      =   6908265
      BevelWidth      =   2
      ScaleWidth      =   12915
      ScaleHeight     =   6495
      Begin VB.TextBox Text1 
         Height          =   2340
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   3405
         Width           =   12555
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   540
         Top             =   5940
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdGeneraDoc 
         Caption         =   "Genera Articoli"
         Height          =   495
         Left            =   10800
         TabIndex        =   8
         Top             =   5820
         Width           =   1920
      End
      Begin FPSpreadADO.fpSpread ssGiacenze 
         Height          =   3045
         Left            =   60
         TabIndex        =   7
         Top             =   165
         Visible         =   0   'False
         Width           =   12705
         _Version        =   524288
         _ExtentX        =   22410
         _ExtentY        =   5371
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DAutoCellTypes  =   0   'False
         DAutoFill       =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         EditEnterAction =   5
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
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
         MaxRows         =   20
         NoBeep          =   -1  'True
         Protect         =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "GeneraCodici.ctx":0425
         UnitType        =   2
         VisibleCols     =   13
         VisibleRows     =   18
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "GeneraCodici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'==============================
'   definizione costanti
'==============================
Const SCH_FILTRO = 0
Const SCH_TABELLA = 1

Private WithEvents frmContenitore As Form
Attribute frmContenitore.VB_VarHelpID = -1
Private FM98 As MXInterfacce.IFunzioniM98
Public WithEvents MWAgt1 As MXKit.CAgenteAuto
Attribute MWAgt1.VB_VarHelpID = -1
Dim WithEvents TabArt As MXKit.CTabelle
Attribute TabArt.VB_VarHelpID = -1
Public WithEvents objFiltro As MXKit.CFiltro
Attribute objFiltro.VB_VarHelpID = -1
Public strNomeFiltro As String

Dim Maschera&
Dim sngLarghSSDesign As Single

Dim ImpComune As Boolean   'Indica se l'impostazione caricata è comune per tutti gli utenti (Usato per l'eventuale eliminazione)

Dim mIntSchOnTop As Integer
Dim mLngButtonMask As Long
Dim MintAccessi As Integer
Dim mVetAgtTab(0 To 1) As String
Dim mBolFiltro As Boolean

Public Property Get Controls(Optional key As Variant) As Object
    On Error Resume Next
    If IsMissing(key) Then
        Set Controls = UserControl.Controls
    Else
        If key <> "" Then
            Set Controls = UserControl.Controls(key)
        End If
    End If
    On Error GoTo 0
End Property

Public Function Inizializza(pfrmCont As Object, colAmbienti As Collection, colOggettiGlobali As Collection, Optional objPar As Object) As Boolean

    Dim vntobj As Variant

    Set frmContenitore = pfrmCont
    
    Set FM98 = pfrmCont.FunzioniM98
    
    Inizializza = Inizializza_i(colAmbienti, colOggettiGlobali)
    
    strNomeFiltro = "ZS_GENERAARTICOLI"
    
    Call Form_Load
    frmContenitore.Caption = "Generazione articoli con variante imballo"

End Function

Public Sub Termina()
    Dim bolCanc As Integer

    Call Form_Unload(bolCanc)
    
    Set FM98 = Nothing
    Set frmContenitore = Nothing
    
    Call Termina_i

End Sub

Private Sub Form_Activate()

    Call MXNU.Attiva_Toolbar(frmContenitore.hwnd, mLngButtonMask)
    Call MXNU.ImpostaFormAttiva(frmContenitore)

End Sub

Private Sub Form_Load()


    Ling(1).Enabled = False
    mLngButtonMask = BTN_STP_MASK
    MintAccessi = FormImpostaAccessi(frmContenitore, mLngButtonMask)
    If (Not InitClassi()) Then
        Call MXNU.MsgBoxEX(1894, vbCritical, 1007)
    Else
        'inizializzo la form
        Call Inizializza_Schede(Me, 2)
        mIntSchOnTop = 1
        Call Ling_GotFocus(SCH_FILTRO)
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim strSQL As String

    Call objFiltro.ImpostaSTSColSS(MXKit.stsSalva)
    
    Set MWAgt1 = Nothing
    
    Set ctlImp.objFiltro = Nothing
    Set ctlImp.cmbListaStp = Nothing
    Set ctlImp.FoglioFiltro = Nothing
    
    Set objFiltro = Nothing
    
    Call TabArt.TChiudiTabella(True)
    Call TabArt.TSalvaImpostazioni
    Set TabArt = Nothing
    
End Sub

Function InitClassi() As Boolean
    Dim strNomeDef As String
    Dim strNomeVis As String

    InitClassi = True
    
    'inizializzo la classe per gli agenti
    Set MWAgt1 = MXAA.CreaCAgenteAuto()
    Call MXAA.RegistraEventiFrm(frmContenitore, MWAgt1)
    
    sngLarghSSDesign = ssFiltro.Width
    Set objFiltro = MXFT.CreaCFiltro()
    mBolFiltro = objFiltro.InizializzaFiltro(strNomeFiltro, ssFiltro)
    If mBolFiltro Then
        
        ssFiltro.VisibleCols = 6
        
        If Not ctlImp.Inizializza(MXDB, MXNU, MXVI, strNomeFiltro, objFiltro, Nothing, ssFiltro, hndDBArchivi) Then
            ctlImp.Visible = False
        End If
        
        If ssFiltro.Width > sngLarghSSDesign Then ssFiltro.Width = sngLarghSSDesign
    Else
        MXNU.MsgBoxEX 1110, vbExclamation, 1007
        InitClassi = False
        Exit Function
    End If
    
    
    Set TabArt = MXCT.CreaCTabelle
    With TabArt
        .NomeTabella = "ZS_GENERAARTICOLI"
        If (.Inizializza(ssGiacenze, Nothing)) Then
            .strWHEAgg = "1=0"
            Call .TApriTabella(True, False)
            Call .TChiudiTabella
            Call .TApriTabella(False)
            .ControlloAggiuntivoTabella = True
            
            Call ssSpreadImposta(ssGiacenze, , , , SS_OP_MODE_ROWMODE)
        End If
    End With

End Function


Private Sub cmdGeneraDoc_Click()
    Dim vet() As String
    Dim cSql As String
    Dim rSql As CRecordSet
    Dim nIdTesta As Long
    Dim vField As ADODB.Field
    Dim i As Long
    Dim nCol As Long
    
    Dim vecchiocodart As String
    Dim nuovocodart As String
    Dim codiceimballo As String

    Text1.Text = "Inizio proceura generazione articoli"
    Text1.Text = Text1.Text & vbCrLf & "------------------------------------------------------------------"
    
        cSql = "SELECT * FROM ZS_GENERAARTICOLI"
        Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)
        i = 1
        Do While Not MXDB.dbFineTab(rSql)
        
            vecchiocodart = MXDB.dbGetCampo(rSql, TIPO_SNAPSHOT, "CODART", "")
            nuovocodart = MXDB.dbGetCampo(rSql, TIPO_SNAPSHOT, "NUOVOCODART", "")
            codiceimballo = MXDB.dbGetCampo(rSql, TIPO_SNAPSHOT, "CODICE", "")
            
            ' OPERAZIONI PRE GENERAZIONE ARTICOLO
            cSql = "EXEC ZS_GENERAARTICOLO_PRE :CODART"
            cSql = Replace(cSql, ":CODART", hndDBArchivi.FormatoSQL(nuovocodart, DB_TEXT))
            Call MXDB.dbEseguiSQL(hndDBArchivi, cSql)
            
            If (GeneraArticolo(nuovocodart) = False) Then
                Call MsgBox("Errore!")
                Exit Sub
            Else
            
                ' OPERAZIONI POST GENERAZIONE ARTICOLO
                cSql = "EXEC ZS_GENERAARTICOLO_POST :CODART, :OLDCODART, :CODIMBALLO"
                cSql = Replace(cSql, ":CODART", hndDBArchivi.FormatoSQL(nuovocodart, DB_TEXT))
                cSql = Replace(cSql, ":OLDCODART", hndDBArchivi.FormatoSQL(vecchiocodart, DB_TEXT))
                cSql = Replace(cSql, ":CODIMBALLO", hndDBArchivi.FormatoSQL(codiceimballo, DB_TEXT))
                
                Text1.Text = Text1.Text & vbCrLf & "Aggiornamento post generazione Articolo: " & nuovocodart & cSql
                Call MXDB.dbEseguiSQL(hndDBArchivi, cSql)
                
                Text1.Text = Text1.Text & vbCrLf & "Aggiornamento post generazione Articolo: " & nuovocodart & " completato!"
            
            End If

            DoEvents

            i = i + 1

            Call MXDB.dbSuccessivo(rSql)
        Loop

        Call MXDB.dbChiudiSS(rSql)
    
        Text1.Text = Text1.Text & vbCrLf & "Generazione articoli terminata!"
        Call MsgBox("Generazione articoli terminata!")
        
    

End Sub

Private Function GeneraArticolo(ByVal strArticolo As String) As Boolean

On Error GoTo err_mng
    
    Dim cSql As String
    Dim rSql As CRecordSet
    Dim bOK As Boolean


    bOK = True
    
    cSql = "SELECT top 1 1 FROM ANAGRAFICAARTICOLI WHERE CODICE = :CODART"
    cSql = Replace(cSql, ":CODART", hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT))
    Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)

    If Not MXDB.dbFineTab(rSql) Then
        
        GeneraArticolo = True
        Text1.Text = Text1.Text & vbCrLf & "Articolo: " & strArticolo & " già generato!"
        Call MXDB.dbChiudiSS(rSql)
        Set rSql = Nothing
        Exit Function

    End If
    
    Call MXDB.dbChiudiSS(rSql)
    Set rSql = Nothing
    
    
    Call MXNU.MostraMsgInfo("Creazione articolo: " & strArticolo)
    
    Dim xCodArt As MXBusiness.CVArt
    Set xCodArt = MXART.CreaCVArt()
    xCodArt.Codice = strArticolo
    If xCodArt.Valida(CHIEDIVAR_NESSUNA, False) Then
        'xCodArt.Descrizione = Descrizione
        Call xCodArt.Genera
        Text1.Text = Text1.Text & vbCrLf & "Generazione articolo: " & strArticolo & " completata!"
    End If
    Set xCodArt = Nothing

    Call MXNU.MostraMsgInfo("")
    
    GeneraArticolo = bOK
    
err_mng:
    If Err.Number <> 0 Then
        Text1.Text = Text1.Text & vbCrLf & Err.Number & " " & Err.Description
        GeneraArticolo = False
    End If

End Function

Private Sub ComProcedi_Click(Index As Integer)
    Call AggTab
    
    Scheda(1).Visible = True
    Scheda(0).Visible = False
    Scheda(1).ZOrder 0
    mIntSchOnTop = 1
End Sub

Private Sub AggTab()
    Dim cSql As String
    Dim hDYSrc As CRecordSet
    Dim NRighe As Long
    Dim bolEnd As Boolean
    
    Call SettaPuntatore(vbHourglass)

    cSql = "DELETE ZS_GENERAARTICOLI "
    Call MXDB.dbEseguiSQL(hndDBArchivi, cSql)
    
    cSql = "INSERT INTO ZS_GENERAARTICOLI (ARTTIPOLOGIA, CODART, CODICE, DESCRIZIONE, VARIANTEIMBALLO, NUOVOCODART, UtenteModifica, DataModifica)"
    cSql = cSql & "SELECT ARTTIPOLOGIA, CODART, CODICE, DESCRIZIONE, varianteimballo, NUOVOCODART, :UTM, :DTM "
    cSql = cSql & "FROM ZS_VISTA_GENERAARTICOLI WHERE :FLT "
    
    cSql = Replace(cSql, ":FLT", objFiltro.SQLFiltro)
    cSql = Replace(cSql, ":I", hndDBArchivi.FormatoSQL(MXNU.IDSessione, DB_DECIMAL))
    cSql = Replace(cSql, ":UTM", hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT))
    cSql = Replace(cSql, ":DTM", hndDBArchivi.FormatoSQL(Now, DB_DATE))
    
    Call MXDB.dbEseguiSQL(hndDBArchivi, cSql)

    With TabArt
        If (.TTabAperta) Then
            Call .TChiudiTabella
        End If
        .strWHEAgg = "1=1"
        Call .TApriTabella(False)
    End With
        
    
    Call SettaPuntatore(vbDefault)
    

End Sub

Private Sub objFiltro_ValidazionePers(ByVal strNomeValid As String, ByVal strNomeCmpValid As String, bolEseguiValStd As Boolean, vntNewValore As Variant)
    Call ValidPersFiltri(strNomeValid, strNomeCmpValid, bolEseguiValStd, vntNewValore)
End Sub


Public Sub MetRegistra()


End Sub


Private Sub frmContenitore_Activate()

    Call Form_Activate
    
End Sub

Private Sub frmContenitore_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Form_QueryUnload(Cancel, UnloadMode)

End Sub

Private Sub Ling_GotFocus(Index As Integer)
    Dim oldling As Integer

    If (Index <> mIntSchOnTop) Then
        oldling = mIntSchOnTop
        mIntSchOnTop = Index
        DoEvents
        If ImpostaScheda Then
            Scheda(Index).Visible = True
            Scheda(oldling).Visible = False
            Ling(oldling).OnTop = False
            Scheda(Index).ZOrder 0
            Ling(Index).OnTop = True
            'Call SendKeys("{TAB}")
        Else
            mIntSchOnTop = oldling
        End If
    End If

End Sub

Public Function AzioniMetodo(setAzione As enmFunzioniMetodo, Optional varparametro As Variant) As Variant

    Select Case setAzione
        Case MetFInserisci
        Case MetFRegistra
        Case MetFAnnulla
        Case MetFPrecedente
        Case MetFSuccessivo
        Case MetFPrimo
        Case MetFUltimo
        Case MetFDettagli
        Case MetFStampa
        Case MetFVisUtenteModifica
        Case MetFMostraCampiDBAnagr
        Case MetFVisDipendenze
        Case Else
    End Select

End Function

Private Sub Scheda_Paint(Index As Integer)

    Call SchedaOmbreggiaControlli(Scheda(Index))

End Sub

Function ImpostaScheda() As Boolean
    Static SbolImpostaScheda As Boolean

    ImpostaScheda = True
    If Not SbolImpostaScheda Then
        SbolImpostaScheda = True
        Select Case mIntSchOnTop
            Case SCH_FILTRO
            
            Case SCH_TABELLA
                
        End Select
        SbolImpostaScheda = False
    End If

End Function

Public Sub SettaPuntatore(intPuntatore As Integer)
    
    MXNU.FrmMetodo.MousePointer = intPuntatore
    
End Sub

Private Sub TabArt_Registrazione(ByVal enmTipoReg As MXKit.setTipoRegistrazione, bolSuccesso As Boolean)
'    Dim i As Long
'    Dim cSql As String
'    Dim nColIdTesta As Long
'    Dim nColIdRiga As Long
'    Dim nIdTesta As Long
'    Dim nIdRiga As Long
'    Dim rSql As CRecordSet
'
'    If enmTipoReg = vePrimaRegistrazione Then
'
'        nColIdTesta = TabArt.TTrovaColonna("PROGRESSIVO")
'        nColIdRiga = TabArt.TTrovaColonna("IDRIGA")
'
'        For i = 1 To ssGiacenze.DataRowCnt
'
'            If ssCellGetValue(ssGiacenze, 0, i) = "M" Then
'
'                nIdTesta = ssCellGetValue(ssGiacenze, nColIdTesta, i)
'                nIdRiga = ssCellGetValue(ssGiacenze, nColIdRiga, i)
'
'                cSql = "SELECT NOTECLI FROM EXTRARIGHEDOC WHERE IDTESTA = :T AND IDRIGA = :R AND ISNULL(NOTECLI,'') <> :N"
'                cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("NOTECLI"), i), DB_TEXT))
'
'                Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)
'
'                If Not MXDB.dbFineTab(rSql) Then
'                    cSql = "UPDATE EXTRARIGHEDOC SET NOTECLI = :N WHERE IDTESTA = :T AND IDRIGA = :R"
'                    cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                    cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                    cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("NOTECLI"), i), DB_TEXT))
'
'                    MXDB.dbEseguiSQL hndDBArchivi, cSql
'
'                End If
'
'                Call MXDB.dbChiudiSS(rSql)
'
'
'                cSql = "SELECT NOTEART FROM EXTRARIGHEDOC WHERE IDTESTA = :T AND IDRIGA = :R AND ISNULL(NOTEART,'') <> :N"
'                cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("NOTEART"), i), DB_TEXT))
'
'                Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)
'
'                If Not MXDB.dbFineTab(rSql) Then
'                    cSql = "UPDATE EXTRARIGHEDOC SET NOTEART = :N WHERE IDTESTA = :T AND IDRIGA = :R"
'                    cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                    cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                    cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("NOTEART"), i), DB_TEXT))
'
'                    MXDB.dbEseguiSQL hndDBArchivi, cSql
'
'                End If
'
'                Call MXDB.dbChiudiSS(rSql)
'
'
'                cSql = "SELECT NOTEMAG FROM EXTRARIGHEDOC WHERE IDTESTA = :T AND IDRIGA = :R AND ISNULL(NOTEMAG,'') <> :N"
'                cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("NOTEMAG"), i), DB_TEXT))
'
'                Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)
'
'                If Not MXDB.dbFineTab(rSql) Then
'                    cSql = "UPDATE EXTRARIGHEDOC SET NOTEMAG = :N WHERE IDTESTA = :T AND IDRIGA = :R"
'                    cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                    cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                    cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("NOTEMAG"), i), DB_TEXT))
'
'                    MXDB.dbEseguiSQL hndDBArchivi, cSql
'
'                End If
'
'                Call MXDB.dbChiudiSS(rSql)
'
'
'                cSql = "SELECT POSIZIONAMENTO FROM EXTRARIGHEDOC WHERE IDTESTA = :T AND IDRIGA = :R AND ISNULL(POSIZIONAMENTO,'') <> :N"
'                cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("POSIZIONAMENTO"), i), DB_TEXT))
'
'                Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)
'
'                If Not MXDB.dbFineTab(rSql) Then
'                    cSql = "UPDATE EXTRARIGHEDOC SET POSIZIONAMENTO = :N WHERE IDTESTA = :T AND IDRIGA = :R"
'                    cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                    cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                    cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("POSIZIONAMENTO"), i), DB_TEXT))
'
'                    MXDB.dbEseguiSQL hndDBArchivi, cSql
'
'                End If
'
'                Call MXDB.dbChiudiSS(rSql)
'
'
'                cSql = "SELECT CONFEZIONATO FROM EXTRARIGHEDOC WHERE IDTESTA = :T AND IDRIGA = :R AND ISNULL(CONFEZIONATO,0) <> :N"
'                cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("CONFEZIONATO"), i), DB_INTEGER))
'
'                Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)
'
'                If Not MXDB.dbFineTab(rSql) Then
'                    cSql = "UPDATE EXTRARIGHEDOC SET CONFEZIONATO = :N WHERE IDTESTA = :T AND IDRIGA = :R"
'                    cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                    cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                    cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("CONFEZIONATO"), i), DB_INTEGER))
'
'                    MXDB.dbEseguiSQL hndDBArchivi, cSql
'
'                End If
'
'                Call MXDB.dbChiudiSS(rSql)
'
'                cSql = "SELECT DISPONIBILITA FROM EXTRARIGHEDOC WHERE IDTESTA = :T AND IDRIGA = :R AND ISNULL(DISPONIBILITA,0) <> :N"
'                cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("DISPONIBILITA"), i), DB_TEXT))
'
'                Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)
'
'                If Not MXDB.dbFineTab(rSql) Then
'                    cSql = "UPDATE EXTRARIGHEDOC SET DISPONIBILITA = :N WHERE IDTESTA = :T AND IDRIGA = :R"
'                    cSql = Replace(cSql, ":T", hndDBArchivi.FormatoSQL(nIdTesta, DB_LONG))
'                    cSql = Replace(cSql, ":R", hndDBArchivi.FormatoSQL(nIdRiga, DB_LONG))
'                    cSql = Replace(cSql, ":N", hndDBArchivi.FormatoSQL(ssCellGetValue(ssGiacenze, TabArt.TTrovaColonna("DISPONIBILITA"), i), DB_DECIMAL))
'
'                    MXDB.dbEseguiSQL hndDBArchivi, cSql
'
'                End If
'
'                Call MXDB.dbChiudiSS(rSql)
'
'
'            End If
'
'        Next
'
'    End If
End Sub


