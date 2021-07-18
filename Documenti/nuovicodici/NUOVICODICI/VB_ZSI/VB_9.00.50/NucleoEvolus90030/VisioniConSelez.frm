VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{DC77D447-8539-4F52-910A-664CAEC7457B}#1.0#0"; "MXKIT.OCX"
Object = "{3A017254-DC53-4E2A-8ACA-DA3674F5DEC8}#1.0#0"; "MXCTRL.OCX"
Begin VB.Form frmVisioniConSelez 
   Appearance      =   0  'Flat
   ClientHeight    =   6360
   ClientLeft      =   1695
   ClientTop       =   2310
   ClientWidth     =   11295
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
   Icon            =   "VisioniConSelez.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
   ScaleWidth      =   11295
   Begin VB.CommandButton com 
      Caption         =   "A&vanti >>"
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
      Index           =   0
      Left            =   9540
      TabIndex        =   7
      Top             =   5880
      WhatsThisHelpID =   25035
      Width           =   1575
   End
   Begin VB.CommandButton com 
      Caption         =   "&Annulla"
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
      Index           =   1
      Left            =   7860
      TabIndex        =   6
      Top             =   5880
      WhatsThisHelpID =   25038
      Width           =   1575
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6360
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
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
      LightColor      =   6908265
      BevelWidth      =   2
      ScaleWidth      =   11295
      ScaleHeight     =   6360
      Begin MXCtrl.MWSchedaBox SchedaTrovaBox 
         Height          =   1155
         Left            =   0
         TabIndex        =   10
         Top             =   4620
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   2037
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LightColor      =   6908265
         ScaleWidth      =   11235
         ScaleHeight     =   1155
         Begin VB.Frame Frame 
            Appearance      =   0  'Flat
            Caption         =   "Visualizzazione Corrente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Index           =   1
            Left            =   3420
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            WhatsThisHelpID =   24015
            Width           =   4335
         End
         Begin VB.Frame Frame 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   2
            Left            =   6780
            TabIndex        =   18
            Top             =   30
            Width           =   4335
            Begin FPSpreadADO.fpSpread ssOpzioni 
               Height          =   735
               Left            =   180
               TabIndex        =   19
               Top             =   180
               Visible         =   0   'False
               Width           =   4035
               _Version        =   524288
               _ExtentX        =   7117
               _ExtentY        =   1296
               _StockProps     =   64
               DisplayColHeaders=   0   'False
               DisplayRowHeaders=   0   'False
               EditModeReplace =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GridShowHoriz   =   0   'False
               GridShowVert    =   0   'False
               MaxCols         =   1
               MaxRows         =   100
               NoBeep          =   -1  'True
               ScrollBars      =   2
               ShadowColor     =   12632256
               ShadowDark      =   8421504
               ShadowText      =   0
               SpreadDesigner  =   "VisioniConSelez.frx":0442
               VisibleCols     =   1
               VisibleRows     =   3
               AppearanceStyle =   0
            End
         End
         Begin VB.Frame Frame 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Index           =   3
            Left            =   60
            TabIndex        =   11
            Top             =   30
            Width           =   6555
            Begin VB.ListBox lstHelp 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   0
               MultiSelect     =   2  'Extended
               TabIndex        =   16
               Top             =   0
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.CommandButton comTrova 
               Caption         =   "Trova"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5460
               TabIndex        =   15
               Top             =   600
               Width           =   930
            End
            Begin VB.CommandButton comAvanzato 
               Caption         =   "Avanzate"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5460
               TabIndex        =   14
               Top             =   300
               Width           =   930
            End
            Begin VB.TextBox TxtTrova 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               HideSelection   =   0   'False
               Left            =   120
               TabIndex        =   13
               Top             =   300
               Width           =   5115
            End
            Begin VB.TextBox txtAvanzato 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               HideSelection   =   0   'False
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Top             =   300
               Visible         =   0   'False
               Width           =   5115
            End
            Begin FPSpreadADO.fpSpread ssTrova 
               Height          =   975
               Left            =   120
               TabIndex        =   17
               Top             =   1020
               Visible         =   0   'False
               Width           =   5130
               _Version        =   524288
               _ExtentX        =   9075
               _ExtentY        =   1720
               _StockProps     =   64
               AutoSize        =   -1  'True
               DisplayRowHeaders=   0   'False
               EditEnterAction =   5
               EditModePermanent=   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   6
               NoBeep          =   -1  'True
               RestrictRows    =   -1  'True
               ScrollBarExtMode=   -1  'True
               ScrollBars      =   2
               SelectBlockOptions=   0
               ShadowColor     =   12632256
               ShadowDark      =   8421504
               ShadowText      =   0
               SpreadDesigner  =   "VisioniConSelez.frx":13D3
               UserResize      =   0
               VisibleCols     =   6
               VisibleRows     =   3
               AppearanceStyle =   0
            End
         End
      End
      Begin VB.ComboBox cmbVisione 
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
         Left            =   6900
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   40
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4515
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   11115
         Begin FPSpreadADO.fpSpread ssVisione 
            DragIcon        =   "VisioniConSelez.frx":2C01
            Height          =   4080
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   10875
            _Version        =   524288
            _ExtentX        =   19182
            _ExtentY        =   7197
            _StockProps     =   64
            AutoCalc        =   0   'False
            AutoClipboard   =   0   'False
            BackColorStyle  =   1
            ColHeaderDisplay=   0
            DAutoHeadings   =   0   'False
            DAutoSave       =   0   'False
            DAutoSizeCols   =   0
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   8421504
            GridColor       =   8421504
            GridShowHoriz   =   0   'False
            MaxCols         =   100
            MaxRows         =   1000000
            MoveActiveOnFocus=   -1  'True
            NoBeep          =   -1  'True
            OperationMode   =   3
            RowHeaderDisplay=   0
            SelectBlockOptions=   0
            ShadowColor     =   12632256
            ShadowDark      =   8421504
            ShadowText      =   0
            SpreadDesigner  =   "VisioniConSelez.frx":2F0B
            UnitType        =   2
            UserResize      =   0
            VirtualOverlap  =   15
            VirtualRows     =   15
            VisibleCols     =   10
            VisibleRows     =   11
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   1
            AppearanceStyle =   0
         End
      End
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6360
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
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
      LightColor      =   6908265
      BevelWidth      =   2
      ScaleWidth      =   11295
      ScaleHeight     =   6360
      Begin VB.ComboBox cmbSelFiltro 
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
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   5880
         Width           =   3135
      End
      Begin MXKit.ctlImpostazioni ctlImpFiltro 
         Height          =   555
         Left            =   1920
         TabIndex        =   4
         Top             =   60
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   979
      End
      Begin FPSpreadADO.fpSpread ssFiltroDati 
         Height          =   4935
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   10995
         _Version        =   524288
         _ExtentX        =   19394
         _ExtentY        =   8705
         _StockProps     =   64
         EditEnterAction =   1
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
         SpreadDesigner  =   "VisioniConSelez.frx":3405
         UnitType        =   2
         AppearanceStyle =   0
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   0
         Left            =   480
         Top             =   5880
         WhatsThisHelpID =   11196
         Width           =   2055
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
         Caption         =   "Selezione Filtro"
      End
   End
End
Attribute VB_Name = "frmVisioniConSelez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1

'======================
'       costanti
'======================
Const SchFiltro = 0
Const SchVisione = 1

Const ComAvanti = 0
Const ComIndietro = 1

'======================
'   classi
'======================
Dim xVisione As MXKit.cTraccia
Dim xInterfaccia As MXKit.cInterfaccia
Public WithEvents xFiltro As MXKit.CFiltro
Attribute xFiltro.VB_VarHelpID = -1
Public WithEvents MWAgt1 As MXKit.CAgenteAuto
Attribute MWAgt1.VB_VarHelpID = -1
Public FormProp As New CFormProp

'======================
'       variabili
'======================
'variabili di stato della form
Dim mIntSchOnTop As Integer
Dim MlngButtonMask As Long
Dim mVetAgtTab(0 To 1) As String
'variabili per visione
Dim mStrNomeVis As String 'nome visione
Dim mStrOrdinamento As String 'datafield colonna ordinamento
Dim mStrCriterio As String 'criterio visione primo livello
Dim mBolInizializza As Boolean  'risultato dell'inizializzazione
Dim mBolFiltro As Boolean  'la visione ha filtro si/no
Dim mStrLogFile As String
Dim mBolImpDefault As Boolean

Private objDaEseguire As Object  '<<< Usato anche da VisioniConSelez_Ext
'Private objDaEseguire As CVisioniConSelez
Private mStrNomeFunzione As String
Private mSngMinWidth As Single
Private mSngMinHeight As Single
Private mSngCurrentWidth As Single
Private mSngCurrentHeight As Single

Public Sub Frm_SetOriginalSize()
    Me.Height = mSngMinHeight
    Me.width = mSngMinWidth
End Sub

' NOME          : VisioneConSelez_Ext
' DESCRIZIONE   : funzione di visione
' PARAMETRO 1   : Classe da eseguire con la collection dei dati selezionati..
' PARAMETRO 2   : nome della visione da rintracciare nel file INI
' PARAMETRO 3   : colonna di ordinamento
' PARAMETRO 4   : stringa SQL che indica il criterio WHERE per il primo livello
Public Function VisioneConSelez_Ext(objClasse As Object, _
                                Optional strNomeVis As Variant, _
                                Optional strColOrdina As Variant, _
                                Optional strCriterio As Variant) As Boolean
                                
    Dim hWndVis As Long

    Set xVisione = MXVI.CreaCTraccia
    Set xInterfaccia = MXVI.CreaCInterfaccia
    VisioneConSelez_Ext = True
    ' associo la classe da eseguire alla fine della selezione..
    Set objDaEseguire = objClasse
On Error GoTo err_InitVisione
    'imposto valori default
    If (IsMissing(strNomeVis)) Then strNomeVis = mStrNomeVis
    If (IsMissing(strColOrdina)) Then strColOrdina = mStrOrdinamento
    If (IsMissing(strCriterio)) Then strCriterio = mStrCriterio
    'imposto i parametri di selezione
    mStrNomeVis = strNomeVis
    mStrOrdinamento = strColOrdina
    mStrCriterio = strCriterio
    'creo le classi di visione
On Error GoTo err_Visione
    'se ho impostato il combo di selezione della visione, allora lo rendo visibile..
    If cmbSelFiltro.ListCount > 0 Then
        cmbSelFiltro.Visible = True
        etc(0).Visible = True
        mStrNomeVis = objDaEseguire.CambiaVisione(cmbSelFiltro.listIndex)
    Else
        cmbSelFiltro.Visible = False
        etc(0).Visible = False
    End If
    'mostro la videata di visione
    If (LoadVisione()) Then
        'imposto le schede
        Call Inizializza_Schede(Me, Scheda.Count)
        mIntSchOnTop = SchVisione
        If (mBolFiltro) Then
            pSchedaOnTop = SchFiltro
        Else
            pSchedaOnTop = SchVisione
        End If
        'imposto il combo di selezione del filtro..
        If cmbSelFiltro.ListCount > 0 Then
            cmbSelFiltro.listIndex = 0
        End If
        'mostro la finestra
        Call CentraFinestra(Me.hwnd)
        Call Me.Show
        'Se premuto Shift non mando l'Avanti in automatico
        If mBolImpDefault And Not (GetAsyncKeyState(VK_LSHIFT) < 0 Or GetAsyncKeyState(VK_RSHIFT) < 0) Then com_Click ComAvanti
    Else
        Call Unload(Me)
    End If

fine_Visione:
    On Local Error GoTo 0
    Exit Function

err_InitVisione:
    VisioneConSelez_Ext = False
    Call MXNU.MsgBoxEX(1054, vbCritical, 1007, Array("Visione", Err.Number, Err.Description))
Resume fine_Visione

err_Visione:
    VisioneConSelez_Ext = False
    Call MXNU.MsgBoxEX(1054, vbCritical, 1007, Array("Visione", Err.Number, Err.Description))
    Call Unload(Me)
Resume fine_Visione
                                
                                
End Function

Private Sub cmbVisione_Click()
    Call xInterfaccia.ListaClick(xVisione, 0)
End Sub


Private Sub ctlImpFiltro_ImpostazioneDefaultCaricata()
    'Rif. Sviluppo Nr. 1041
    mBolImpDefault = True
End Sub


'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
'           EVENTI OGGETTI DELLA FORM
'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
Private Sub Form_Activate()
    Call MXNU.Attiva_Toolbar(hwnd, MlngButtonMask)
    Call MXNU.ImpostaFormAttiva(Me)
End Sub

Private Sub Form_Load()
    Dim strDim As String
    
    MlngButtonMask = BTN_STP_MASK
    Me.HelpContextID = FormProp.FormID
    Set MWAgt1 = MXAA.CreaCAgenteAuto()
    Call MXAA.RegistraEventiFrm(Me, MWAgt1)
    Call MWAgt1.RegistraAgenteFrm(Me)
    Call MXNU.LeggiRisorseControlli(Me)
    'ATTENZIONE: NON DECOMMENTARE QUESTA RIGA ALTRIMENTI NON FUNZIONA PIU'
    '            LA CHIMAMATA DI CLASSI CON FUNZIONI TARGET DIVERSE DA USACOLLECTION
    'mStrNomeFunzione = ""
    
    'Leggo le dimensioni originali della form dal file MWForm.Ini
    strDim = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MWForm.ini", Me.Name, "Height", "")
    If strDim <> "" Then
        Me.Height = Val(strDim)
    End If
    strDim = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MWForm.ini", Me.Name, "Width", "")
    If strDim <> "" Then
        Me.width = Val(strDim)
    End If
    '...e le memorizzo
    mSngMinHeight = Me.Height
    mSngMinWidth = Me.width
    mSngCurrentHeight = mSngMinHeight
    mSngCurrentWidth = mSngMinWidth
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call xFiltro.ImpostaSTSColSS(stsSalva)
    'Per Metodo Evolus
'    If Not Cancel Then
'        On Local Error Resume Next
'        If (Not mResize Is Nothing) Then
'                mResize.Terminate
'                Set mResize = Nothing
'        End If
'        On Local Error GoTo 0
'    End If
End Sub

Private Sub Form_Resize()
Dim sngWidth As Single
Dim sngHeight As Single
Dim sngDeltaWidth As Single
Dim sngDeltaHeight As Single
Dim i As Integer

    On Local Error Resume Next
    If (Not Me.Visible) Or (Me.WindowState = vbMinimized) Then Exit Sub

    If (Me.WindowState = vbNormal) Then
        If (Me.width < mSngMinWidth Or Me.Height < mSngMinHeight) Then
            Me.width = mSngMinWidth
            Me.Height = mSngMinHeight
        End If
    End If
    sngWidth = Me.width
    sngDeltaWidth = sngWidth - mSngCurrentWidth
    mSngCurrentWidth = sngWidth

    sngHeight = Me.Height
    sngDeltaHeight = sngHeight - mSngCurrentHeight
    mSngCurrentHeight = sngHeight

    For i = 0 To 1
        Scheda(i).Height = Scheda(i).Height + sngDeltaHeight
        Scheda(i).width = Scheda(i).width + sngDeltaWidth
    Next i
    
    ssFiltroDati.Height = ssFiltroDati.Height + sngDeltaHeight
    ssFiltroDati.width = ssFiltroDati.width + sngDeltaWidth
    
    ssVisione.Height = ssVisione.Height + sngDeltaHeight
    ssVisione.width = ssVisione.width + sngDeltaWidth
    cmbVisione.Left = cmbVisione.Left + sngDeltaWidth
    If MXNU.ResizeProporzionale Then
        If (Me.Height - mSngMinHeight) > (mSngMinHeight * MXNU.PercResizeProporzionale / 100) Then ssSpreadSetFontSize ssVisione, 10 Else ssSpreadSetFontSize ssVisione, 8
    End If
    
    Frame(0).Height = Frame(0).Height + sngDeltaHeight
    Frame(0).width = Frame(0).width + sngDeltaWidth
    SchedaTrovaBox.Top = SchedaTrovaBox.Top + sngDeltaHeight
    SchedaTrovaBox.width = SchedaTrovaBox.width + sngDeltaWidth
    
    com(0).Top = com(0).Top + sngDeltaHeight
    com(0).Left = com(0).Left + sngDeltaWidth
    com(1).Top = com(1).Top + sngDeltaHeight
    com(1).Left = com(1).Left + sngDeltaWidth
    On Local Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = vbNormal Then
        Call MXNU.ScriviProfilo(MXNU.PercorsoPreferenze & "\MWForm" & MXNU.NTerminale & ".ini", "VISIONICONSELEZ", "Width", Me.width)
        Call MXNU.ScriviProfilo(MXNU.PercorsoPreferenze & "\MWForm" & MXNU.NTerminale & ".ini", "VISIONICONSELEZ", "Height", Me.Height)
    End If
    Set MWAgt1 = Nothing
    Call xVisione.TerminaVisione
    Set xVisione = Nothing
    Set xInterfaccia = Nothing
    If mBolFiltro Then
        Call ctlImpFiltro.Termina
    End If
    Set xFiltro = Nothing
    Set FormProp = Nothing
    Set objDaEseguire = Nothing
    Set frmVisioniConSelez = Nothing
End Sub

Private Sub cmbSelFiltro_Click()
    ssFiltroDati.Visible = False
    Set MWAgt1 = Nothing
    Call xVisione.TerminaVisione(False)
    'Call xVisione.TerminaVisione
    Set xVisione = Nothing
    Set xInterfaccia = Nothing
    Set xVisione = MXVI.CreaCTraccia
    Set xInterfaccia = MXVI.CreaCInterfaccia
    mStrNomeVis = objDaEseguire.CambiaVisione(cmbSelFiltro.listIndex)
    Call LoadVisione
    ssFiltroDati.Visible = True
End Sub

Private Sub com_Click(Index As Integer)
Dim bolEsitoPositivo As Boolean

    Select Case Index
        Case ComAvanti
            If (mIntSchOnTop = SchVisione) Then
                If (LeggiRigheSelezionate()) Then
                    If Len(mStrNomeFunzione) = 0 Then
                        metodo.MousePointer = vbHourglass
                        Dim bolInterrotto As Boolean
                        Dim bolEsci As Boolean
                        'funzione standard
                        Call objDaEseguire.UsaCollection(xVisione.colRisultatoSel, cmbSelFiltro.listIndex, xFiltro, ssFiltroDati)
                        On Local Error Resume Next
                        bolInterrotto = objDaEseguire.ElabInterrotta
                        If Err.Number = 0 Then
                            bolEsci = Not bolInterrotto
                        Else
                            bolEsci = True
                        End If
                        On Local Error GoTo 0
                        If bolEsci Then
                            Call MXNU.MsgBoxEX(1639, vbInformation, Me.Caption)
                            Unload Me
                        End If
                        metodo.MousePointer = vbDefault
                    Else
                        metodo.MousePointer = vbHourglass
                        'funzione con nome diverso
                        CallByName objDaEseguire, mStrNomeFunzione, VbMethod, xVisione.colRisultatoSel, cmbSelFiltro.listIndex, xFiltro, ssFiltroDati
                        Call MXNU.MsgBoxEX(1639, vbInformation, Me.Caption)
                        Unload Me
                        metodo.MousePointer = vbDefault
                    End If
                End If
            Else
                pSchedaOnTop = mIntSchOnTop + 1
            End If
        Case ComIndietro
            If (mIntSchOnTop = SchFiltro) Then
                Unload Me
            Else
                pSchedaOnTop = mIntSchOnTop - 1
            End If
    End Select
End Sub

Private Sub Scheda_Paint(Index As Integer)
    Call SchedaOmbreggiaControlli(Scheda(Index))
End Sub

Private Sub ssFiltroDati_Change(ByVal Col As Long, ByVal Row As Long)
    'Sviluppo 1077
    Dim TipoElab As String
    If objDaEseguire.TipoElaborazione = VCS_TRASFDEFINITIVE Then
        If Row = xFiltro.IdFiltro2Riga(7) Then
            Call xFiltro.FireSSChange(Col, Row)
            TipoElab = xFiltro.ParAgg("NewTipo").ValoreFormula
            If TipoElab = "E" Then
                Call xFiltro.AttivaRigheFiltroAgg("E")
            Else
                Call xFiltro.AttivaRigheFiltroAgg("M")
            End If
        End If
    ElseIf objDaEseguire.TipoElaborazione = VCS_GENERACONSUMICLAV Then
        If Row = xFiltro.IdFiltro2Riga(2) Then
            Call xFiltro.FireSSChange(Col, Row)
            TipoElab = xFiltro.ParAgg("TipoElaborazione").ValoreFormula
            If TipoElab = "1" Then
                Call xFiltro.AttivaRigheFiltroAgg("1")
            ElseIf TipoElab = "2" Then
                Call xFiltro.AttivaRigheFiltroAgg("2")
            Else
                Call xFiltro.AttivaRigheFiltroAgg("3")
            End If
        End If
    End If
End Sub


Private Sub ssOpzioni_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Call xInterfaccia.FiltroButtonClicked(xVisione, 1, Col, Row, ButtonDown)
End Sub

Private Sub ssVisione_Click(ByVal Col As Long, ByVal Row As Long)
    Call xInterfaccia.VisioneClick(xVisione, 1, Col, Row)
End Sub

Private Sub ssVisione_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)
    Call xInterfaccia.VisioneColWidthChange(xVisione, 1, Col1)
End Sub

Private Sub ssVisione_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call xInterfaccia.VisioneDblClick(xVisione, 1, Col, Row)
End Sub

Private Sub ssVisione_DragDrop(Source As Control, x As Single, y As Single)
    Call xInterfaccia.VisioneDragDrop(xVisione, 1, Source, x, y)
End Sub

Private Sub ssVisione_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Call xInterfaccia.VisioneDragOver(xVisione, 1, Source, x, y, State)
End Sub

Private Sub ssVisione_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call xInterfaccia.VisioneMouseDown(xVisione, 1, Button, Shift, x, y)
End Sub

Private Sub ssVisione_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call xInterfaccia.VisioneMouseMove(xVisione, 1, Button, Shift, x, y)
End Sub

Private Sub ssVisione_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    Call xInterfaccia.VisioneTopLeftChange(xVisione, 1, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Private Sub xFiltro_ValidazionePers(ByVal strNomeValid As String, ByVal strNomeCmpValid As String, bolEseguiValStd As Boolean, vntNewValore As Variant)
    Call ValidPersFiltri(strNomeValid, strNomeCmpValid, bolEseguiValStd, vntNewValore)
End Sub

'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
'           FUNZIONI PRIVATE DELLA FORM
'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
Private Property Let pSchedaOnTop(new_Valore As Integer)
    Dim bolValido As Boolean
    bolValido = True
    Select Case new_Valore
        Case SchFiltro
            com(ComAvanti).Caption = MXNU.CaricaStringaRes(25035) 'avanti
            com(ComIndietro).Caption = MXNU.CaricaStringaRes(25038) 'annulla
        Case SchVisione
            com(ComAvanti).Caption = MXNU.CaricaStringaRes(25005) 'procedi
            com(ComIndietro).Caption = MXNU.CaricaStringaRes(25036) 'indietro
            com(ComIndietro).Enabled = mBolFiltro
            If (new_Valore > mIntSchOnTop) Then 'carico la visione solo se sto andando avanti
                If xFiltro.CtrlCampiObbligatori() Then   'Rif. Anomalia Nr. 5943: Aggiunto controllo su campi obbligatori del filtro
                    Me.MousePointer = vbHourglass
                    With xVisione
                        If mBolFiltro Then
                            .colLivelli(1).strSQLWhr = xFiltro.SQLFiltro
                            Call .ImpostaSegnaposto
                        End If
                        If (Not .colLivelli(1).bolImpostato) Then
                            'prima volta->imposto livello 1
                            .LivelloCorrente = 1
                            Call .CalcVisibleRows
                        Else
                            'ricarico la visione
                            Call .VisioneCaricaDati(1, .colLivelli(1).mIntVisione)
                        End If
                    End With
                    Me.MousePointer = vbDefault
                Else
                    bolValido = False
                    com(ComAvanti).Caption = MXNU.CaricaStringaRes(25035) 'avanti
                    com(ComIndietro).Caption = MXNU.CaricaStringaRes(25038) 'annulla
                End If
            End If
    End Select
    If bolValido Then
        DoEvents
        Scheda(new_Valore).Visible = True
        Scheda(mIntSchOnTop).Visible = False
        Call Scheda(new_Valore).ZOrder(vbBringToFront)
        Call com(ComAvanti).ZOrder(vbBringToFront)
        Call com(ComIndietro).ZOrder(vbBringToFront)
        mIntSchOnTop = new_Valore
        '******* Remmato altrimenti non funziona l'avvio dell'help con F1 dalla prima scheda.
        'Me.KeyPreview = (mIntSchOnTop = SchVisione)
        '*************************************************************************
    End If
End Property

Private Function LoadVisione() As Boolean
Dim intCurLiv As Integer

    Screen.MousePointer = HOURGLASS

    Call MXNU.LeggiRisorseControlli(Me)
    mBolImpDefault = False
    mBolInizializza = xVisione.Inizializza(mStrNomeVis, MXKit.tivVisione, mStrOrdinamento, mStrCriterio, MXKit.selSelezioneCheck, Nothing)
    If (mBolInizializza) Then
        'imposto il foglio del filtro
        Set xFiltro = MXFT.CreaCFiltro()
        mBolFiltro = (xFiltro.InizializzaFiltro(xVisione.colLivelli(1).strFiltro, ssFiltroDati))
        If objDaEseguire.TipoElaborazione = VCS_TRASFDEFINITIVE Then Call xFiltro.AttivaRigheFiltroAgg("M")
        If objDaEseguire.TipoElaborazione = VCS_GENERACONSUMICLAV Then Call xFiltro.AttivaRigheFiltroAgg("1")
        On Local Error Resume Next
        Set objDaEseguire.objFiltro = xFiltro
        On Local Error GoTo 0
        If (mBolFiltro) Then
            Call ctlImpFiltro.Inizializza(MXDB, MXNU, MXVI, xVisione.colLivelli(1).strFiltro, xFiltro, Nothing, ssFiltroDati, hndDBArchivi, GIMP_FILTRONORMALE)
            Set xVisione.CFiltroDati = xFiltro
        End If
        'imposto i controlli della traccia
        Call xVisione.ImpostaControlli(Me, _
            Nothing, _
            TxtTrova, _
            txtAvanzato, _
            ssTrova, _
            comTrova, _
            comAvanzato, _
            lstHelp, _
            , , , _
            ssFiltroDati)
        'carico i controlli
        Call xVisione.colLivelli(1).ImpostaControlli(ssVisione, cmbVisione, ssOpzioni)
        Screen.MousePointer = vbDefault
    End If
    'risultato
    LoadVisione = mBolInizializza
End Function

Private Function LeggiRigheSelezionate() As Boolean
    Me.MousePointer = vbHourglass
    'resetto la collection dei risultati
    If (xVisione.colRisultatoSel.Count > 0) Then
        Set xVisione.colRisultatoSel = Nothing
        Set xVisione.colRisultatoSel = New Collection
    End If
    'seleziono il progressivo
    Call xVisione.RisultatiSelezione
    LeggiRigheSelezionate = (xVisione.colRisultatoSel.Count > 0)
    If (Not LeggiRigheSelezionate) Then Call MXNU.MsgBoxEX(2049, vbExclamation, 1007)
    Me.MousePointer = vbDefault
End Function

'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
'           FUNZIONI PUBBLICHE DELLA FORM
'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&

' NOME          : Visione
' DESCRIZIONE   : funzione di visione
' PARAMETRO 1   : Tipo di Elaborazione da eseguire con la collection dei dati selezionati..
' PARAMETRO 2   : nome della visione da rintracciare nel file INI
' PARAMETRO 3   : colonna di ordinamento
' PARAMETRO 4   : stringa SQL che indica il criterio WHERE per il primo livello
Public Function VisioneConSelez(TipoElaborazione As setTipoVisioneConSelez, _
                                Optional strNomeVis As Variant, _
                                Optional strColOrdina As Variant, _
                                Optional strCriterio As Variant) As Boolean

    Dim hWndVis As Long
    Dim sngSavedWidth As Single
    Dim sngSavedHeight As Single

    Set xVisione = MXVI.CreaCTraccia
    Set xInterfaccia = MXVI.CreaCInterfaccia
    VisioneConSelez = True
    ' associo la classe da eseguire alla fine della selezione..
    'Set objDaEseguire = objClasse
    Set objDaEseguire = New CVisioniConSelez
    objDaEseguire.TipoElaborazione = TipoElaborazione
On Error GoTo err_InitVisione
    'imposto valori default
    If (IsMissing(strNomeVis)) Then strNomeVis = mStrNomeVis
    If (IsMissing(strColOrdina)) Then strColOrdina = mStrOrdinamento
    If (IsMissing(strCriterio)) Then strCriterio = mStrCriterio
    'imposto i parametri di selezione
    mStrNomeVis = strNomeVis
    mStrOrdinamento = strColOrdina
    mStrCriterio = strCriterio
    'creo le classi di visione
On Error GoTo err_Visione
    'se ho impostato il combo di selezione della visione, allora lo rendo visibile..
    If cmbSelFiltro.ListCount > 0 Then
        cmbSelFiltro.Visible = True
        etc(0).Visible = True
        mStrNomeVis = objDaEseguire.CambiaVisione(cmbSelFiltro.listIndex)
    Else
        cmbSelFiltro.Visible = False
        etc(0).Visible = False
    End If
    'mostro la videata di visione
    If (LoadVisione()) Then
        'imposto le schede
        Call Inizializza_Schede(Me, Scheda.Count)
        mIntSchOnTop = SchVisione
        If (mBolFiltro) Then
            pSchedaOnTop = SchFiltro
        Else
            pSchedaOnTop = SchVisione
        End If
        'imposto il combo di selezione del filtro..
        If cmbSelFiltro.ListCount > 0 Then
            cmbSelFiltro.listIndex = 0
        End If
        'carico le impostazioni
        sngSavedWidth = Cast(MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MWForm" & MXNU.NTerminale & ".ini", "VISIONICONSELEZ", "Width", "0"), vbSingle)
        sngSavedHeight = Cast(MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MWForm" & MXNU.NTerminale & ".ini", "VISIONICONSELEZ", "Height", "0"), vbSingle)
        If (sngSavedWidth <> 0 And sngSavedWidth <> mSngMinWidth And sngSavedHeight <> 0 And sngSavedHeight <> mSngMinHeight) Then
            Me.Move 0, 0, sngSavedWidth, sngSavedHeight
        End If
        
        'mostro la finestra
        Call CentraFinestra(Me.hwnd)
        'Inzializzazione Form per Metodo Evolus
        Call CambiaColoriControlli(Me)   'Da Mettere PRIMA di mResize
        SchedaTrovaBox.ShadowColor = SysGradientColor1
        On Local Error Resume Next
        'Set mResize = New MxResizer.ResizerEngine
        'If (Not mResize Is Nothing) Then
        '        Call mResize.Initialize(Me, , , , , 0, True, metodo, MXNU)
        'End If
        Call CambiaCharSet(Me)
        On Local Error GoTo 0
        Call Me.Show
        'Se premuto Shift non mando l'Avanti in automatico
        If mBolImpDefault And Not (GetAsyncKeyState(VK_LSHIFT) < 0 Or GetAsyncKeyState(VK_RSHIFT) < 0) Then com_Click ComAvanti
    Else
        Call Unload(Me)
    End If

fine_Visione:
    On Local Error GoTo 0
    Exit Function

err_InitVisione:
    VisioneConSelez = False
    Call MXNU.MsgBoxEX(1054, vbCritical, 1007, Array("Visione", Err.Number, Err.Description))
Resume fine_Visione

err_Visione:
    VisioneConSelez = False
    Call MXNU.MsgBoxEX(1054, vbCritical, 1007, Array("Visione", Err.Number, Err.Description))
    Call Unload(Me)
Resume fine_Visione
End Function

Public Function AzioniMetodo(setAzione As enmFunzioniMetodo, Optional varparametro As Variant) As Variant

    Select Case setAzione
        Case MetFSchedulaOperazione
            If (StrComp(mStrNomeFunzione, "RivalutaCommesse", vbTextCompare) = 0) Then
                If ctlImpFiltro.NomeImpostazione <> "" Then
                    Call NuovaSchedulazione
                End If
            End If
        Case MetFInserisci
        
        Case MetFRegistra
        Case MetFAnnulla
        Case MetFPrecedente
        Case MetFSuccessivo
        Case MetFPrimo
        Case MetFUltimo
        Case MetFDettagli
        Case MetFStampa
            Call xVisione.Stampa
        Case MetFVisUtenteModifica
        Case MetFDettVisione
        Case MetFMostraCampiDBAnagr
        Case Else
    End Select

End Function

' modifica del 25/02/2002 - Utilizzato il nuovo job scheduler
Private Function NuovaSchedulazione() As Boolean

#If BATCH <> 1 And ISM98SERVER <> 1 And TOOLS <> 1 And ISNUCLEO = 0 Then

Dim bolRes As Boolean
Dim objSchedula As MxScheduler.clsScheduler
Dim objOperDb As clsOperDb

    bolRes = False
    On Local Error GoTo ERR_NuovaSchedulazione
    Set objSchedula = New MxScheduler.clsScheduler
    If objSchedula.Inizializza(MXNU, Command()) Then
        Set objOperDb = objSchedula.CreaOperazione("RIVALUTACOMMESSE", True)
        objOperDb.Descrizione = ctlImpFiltro.NomeImpostazione
        Call objOperDb.SetRiga("IMPOSTAZIONE FILTRO COMMESSE", 3, ctlImpFiltro.NomeImpostazione)
        Call objSchedula.NewOperation(objOperDb, False)
        bolRes = True
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

'Public Property Let pNomeFuzioneUsaCollection(vData As String)
'    mStrNomeFunzione = vData
'End Property




'Per Metodo Evolus
'Private Sub mResize_AfterResize()
'    Call AvvicinaLing(Me)
'End Sub

