VERSION 5.00
Object = "{4AFA2505-EEFF-4BA2-873D-9FDF23CDB0CB}#1.0#0"; "MXCTRL.OCX"
Begin VB.Form frmGlobale 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anagrafica Globale"
   ClientHeight    =   6675
   ClientLeft      =   1380
   ClientTop       =   2205
   ClientWidth     =   9675
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
   HelpContextID   =   2010
   Icon            =   "FrmGlobale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   645
   ShowInTaskbar   =   0   'False
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      HelpContextID   =   2011
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   21001
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Dati &Anagr."
      First           =   -1  'True
   End
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      HelpContextID   =   2019
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Tag             =   "lingext"
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "ESTENSIONE"
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6330
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   11165
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
      ScaleWidth      =   9675
      ScaleHeight     =   6330
      Begin MXCtrl.MWEtichetta etc 
         Height          =   375
         Index           =   1
         Left            =   240
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
         ForeColor       =   6697728
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Descrizione"
         UseGradientColor=   0   'False
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BackColor       =   -2147483633
         ForeColor       =   6697728
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Codice"
         UseGradientColor=   0   'False
      End
      Begin VB.TextBox txtb 
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
         Index           =   1
         Left            =   2040
         TabIndex        =   2
         Text            =   "Rag. Sociale"
         Top             =   675
         Width           =   7335
      End
      Begin VB.TextBox txtb 
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
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Text            =   "X 99999"
         Top             =   195
         Width           =   915
      End
      Begin MXCtrl.XPToolButton tbsel 
         Height          =   285
         Index           =   0
         Left            =   3000
         Top             =   180
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomButton    =   1
      End
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6330
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   11165
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
      ScaleWidth      =   9675
      ScaleHeight     =   6330
   End
End
Attribute VB_Name = "frmGlobale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'=================================================
'   definizione classi
'=================================================
Dim WithEvents AnaGlobale As MXKit.Anagrafica
Attribute AnaGlobale.VB_VarHelpID = -1

'=================================================
'   definizione classi publiche
'=================================================
Public WithEvents MWAgt1 As MXKit.CAgenteAuto
Attribute MWAgt1.VB_VarHelpID = -1
Public FunzioniM98 As CFunzioniMetodo98
Public FormProp As New CFormProp

'=================================================
'   definizione variabili publiche
'=================================================
Public TipoConto As String

'=================================================
'   definizione variabili private
'=================================================
Dim mIntSchOnTop As Integer
Dim mCtlExt As VBControlExtender
Dim strCaption As String
Dim strTipoEstensione As String

Private Sub Form_Activate()
    Call MXNU.Attiva_Toolbar(hwnd, BTN_TUTTI_MASK)
    Call MXNU.ImpostaFormAttiva(Me)
End Sub

Public Sub MetInserisci()
    On Local Error Resume Next
    Call AnaGlobale.Inserisci
    If Not Scheda(0).Visible Then
        Ling(0).SetFocus
    End If
    txtb(0).SetFocus
    On Local Error GoTo 0
End Sub

Public Function MetRegistra() As Boolean
    MetRegistra = (AnaGlobale.SalvaAnagrafica() = saRegCorretta)
End Function

Sub MetAnnulla()
    If (MXNU.MsgBoxEX(MXNU.CaricaStringaRes(1014, txtb(0).Text), 36, MXNU.CaricaStringaRes(1007)) = vbYes) Then
        Call AnaGlobale.Annulla
    End If
    Call MetInserisci
End Sub

Sub MetPrimo()
    Call AnaGlobale.FrecceSpostamento(BTN_PRIMO)
End Sub

Sub MetPrecedente()
    Call AnaGlobale.FrecceSpostamento(BTN_PREC)
End Sub

Sub MetSuccessivo()
    Call AnaGlobale.FrecceSpostamento(BTN_SUCC)
End Sub

Sub MetUltimo()
    Call AnaGlobale.FrecceSpostamento(BTN_ULTIMO)
End Sub

Private Sub Form_Load()
    
  Dim intRes As Integer
  
    metodo.MousePointer = vbHourglass
      
    Me.HelpContextID = FormProp.FormID
    
    'inizializzazione agenti
    Set MWAgt1 = MXAA.CreaCAgenteAuto()
    intRes = MXAA.RegistraEventiFrm(Me, MWAgt1)


    'Caricamento dell'anagrafica corrispondente alla voce di menù scelta
    Call ImpostaForm(Me.HelpContextID)
    
    'Reset  dei Campi
    Call ResetCampi
    
    ' Assegnamento delle risorse in lingua
    Call DefLingua
       
    AnaGlobale.flgModRecord = False
    
    If (TipoConto = "F") Then
        Call AnaGlobale.McolControlli("txtb_0").Inizializza("VALID_ANAGFOR", True)
    End If
    
    'definizione dell'estensione
    Ling(1).Visible = DefEstensione(strTipoEstensione, Ling(1), Scheda(1), Ling(0).Left + Ling(0).Width, AnaGlobale, mCtlExt, 700)
    
    'visualizzazione della finestra
    Call CentraFinestra(Me.hwnd)
    Me.Show
    
    metodo.MousePointer = vbDefault

End Sub

Private Sub ResetCampi()
    txtb(0).Text = ""
    txtb(1).Text = ""
End Sub

Private Sub ImpostaForm(lngHelpID As Long)
  Dim strNomeAnagrafica As String
  
    Select Case lngHelpID
        Case 1400  ' Clienti
            strNomeAnagrafica = "AnagraficaCF"
            strCaption = MXNU.CaricaStringaRes(23010)
            strTipoEstensione = "CLIENTI"
        Case 1405   ' Fornitori
            strNomeAnagrafica = "AnagraficaCF"
            strCaption = MXNU.CaricaStringaRes(23011)
            strTipoEstensione = "FORNITORI"
        Case 1410   ' Generici
            strNomeAnagrafica = "AnagraficaGenerici"
            strCaption = MXNU.CaricaStringaRes(23012)
            strTipoEstensione = "GENERICI"
        Case 1415   ' Agenti
            strNomeAnagrafica = "AnagraficaAgenti"
            strCaption = MXNU.CaricaStringaRes(23001)
            strTipoEstensione = "AGENTI"
        Case 1420   ' Banche
            strNomeAnagrafica = "AnagraficaBanche"
            strCaption = MXNU.CaricaStringaRes(23002)
            strTipoEstensione = "BANCHE"
        Case 3000   ' Articoli
            strNomeAnagrafica = "AnagraficaArticoli"
            strCaption = MXNU.CaricaStringaRes(23038)
            strTipoEstensione = "MAG"
        Case 3005   ' Articoli con tipologie
            strNomeAnagrafica = "AnagraficaArticoli"
            strCaption = MXNU.CaricaStringaRes(23027)
            strTipoEstensione = "MAG"
        Case 3010   ' Magazzini
            strNomeAnagrafica = "AnagraficaDepositi"
            strCaption = MXNU.CaricaStringaRes(23017)
            strTipoEstensione = "DEPOSITI"
        Case 4005   ' Distinta Base
            strNomeAnagrafica = "DistintaBase"
            strCaption = MXNU.CaricaStringaRes(23079)
            strTipoEstensione = "DISTINTA"
        Case 5516  ' Cespiti
            strNomeAnagrafica = "AnagraficaCespiti"
            strCaption = MXNU.CaricaStringaRes(23249)
            strTipoEstensione = "CESPITI"
    End Select
        
    ' Se cliente o fornitore utilizzo variabile TipoConto
    If lngHelpID = 1400 Or lngHelpID = 1405 Then
        Set AnaGlobale = MXVA.CreaCAnagrafica(strNomeAnagrafica, Me, TipoConto)
    Else
        Set AnaGlobale = MXVA.CreaCAnagrafica(strNomeAnagrafica, Me)
    End If

End Sub

Private Sub DefLingua()
    'Caricamento della caption nella finestra
    Me.Caption = strCaption
    ' Associo le risorse in lingua alle etichette dipendenti
    Call AnaGlobale.Disegna
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If (Not MWAgt1 Is Nothing) Then
        Call MWAgt1.Termina
        Set MWAgt1 = Nothing
    End If
    Call TerminaEstensione(mCtlExt, Scheda(1))
    Set FormProp = Nothing
    Set frmGlobale = Nothing

End Sub

Private Sub Ling_GotFocus(Index As Integer)
    metodo.MousePointer = vbHourglass
    If (Index <> mIntSchOnTop) Then
        Scheda(Index).Visible = True
        
        If (mIntSchOnTop <> -1) Then
            Scheda(mIntSchOnTop).Visible = False
            Ling(mIntSchOnTop).OnTop = False
        End If
        Scheda(Index).ZOrder vbBringToFront
        Ling(Index).OnTop = True
        mIntSchOnTop = Index
    End If
    metodo.MousePointer = vbDefault
End Sub

Private Sub tbSel_Click(Index As Integer)
Dim strRifControllo As String

    On Local Error Resume Next
    strRifControllo = "tbsel_" & Index
    If (AnaGlobale.NomeSel2NomeGruppo(strRifControllo) <> "") Then
        Call AnaGlobale.TBselClick(strRifControllo, "")
    End If
    On Local Error GoTo 0
End Sub

Private Sub txtb_GotFocus(Index As Integer)
    SelContenuto txtb(Index)
End Sub

Private Sub txtb_KeyPress(Index As Integer, keyAscii As Integer)
    Call CtrlKey(keyAscii, AnaGlobale.GrInput("txtb_" & Index).TipoInput)
End Sub

Private Sub txtb_LostFocus(Index As Integer)
    metodo.MousePointer = vbHourglass
    On Local Error Resume Next
    Call AnaGlobale.AssegnaCampo("txtb_" & Index, txtb(Index).Text)
    metodo.MousePointer = vbDefault
End Sub

Public Function AzioniMetodo(setAzione As enmFunzioniMetodo, Optional varparametro As Variant) As Variant
    Dim ListaCol As New Collection
    
    MXNU.MostraMsgInfo ""
    Select Case setAzione
        Case MetFInserisci:  Call MetInserisci
        Case MetFRegistra:   AzioniMetodo = MetRegistra()
        Case MetFAnnulla:    Call MetAnnulla
        Case MetFPrecedente: Call MetPrecedente
        Case MetFSuccessivo: Call MetSuccessivo
        Case MetFPrimo:      Call MetPrimo
        Case MetFUltimo:     Call MetUltimo
        Case MetFDettagli
        Case MetFStampa
         Case MetFVisUtenteModifica
        Case MetFDettVisione
        Case MetFMostraCampiDBAnagr
        Case MetFVisDipendenze
        Case Else
    End Select
    On Local Error Resume Next
    Call mCtlExt.object.AzioniMetodo(setAzione, varparametro)
End Function
