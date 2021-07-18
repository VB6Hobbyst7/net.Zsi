VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A14AEF4A-9F3F-48DA-8192-94EB9D4AAB06}#1.0#0"; "MXCTRL.OCX"
Begin VB.Form frmSelUtente 
   Caption         =   "Selezione Utente"
   ClientHeight    =   1635
   ClientLeft      =   2670
   ClientTop       =   1875
   ClientWidth     =   5535
   ControlBox      =   0   'False
   Icon            =   "SelUtente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5535
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   1635
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleWidth      =   5535
      ScaleHeight     =   1635
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         Caption         =   "Nome Utente/Gruppo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   3495
         Begin MSComctlLib.TreeView trwUtenti 
            Height          =   3735
            Left            =   420
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   300
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   6588
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            PathSeparator   =   ";"
            Style           =   7
            ImageList       =   "ImgLst"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.ImageList ImgLst 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   14
            ImageHeight     =   15
            MaskColor       =   12632256
            UseMaskColor    =   0   'False
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   12
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":000C
                  Key             =   "metodo98"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":05B2
                  Key             =   "entire"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":06C4
                  Key             =   "gruppo"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":0C6A
                  Key             =   "utente"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":1210
                  Key             =   "formab"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":1322
                  Key             =   "formds"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":1874
                  Key             =   "lingab"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":196E
                  Key             =   "lingds"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":1A68
                  Key             =   "sitab"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":1B82
                  Key             =   "sitds"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":1C9C
                  Key             =   "findab"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SelUtente.frx":1DEE
                  Key             =   "findds"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         Caption         =   "Inserire il Gruppo/Utente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   120
         WhatsThisHelpID =   24026
         Width           =   3495
         Begin VB.TextBox txtb 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   3195
         End
      End
      Begin VB.CommandButton com 
         Caption         =   "&Annulla"
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   2
         Top             =   600
         WhatsThisHelpID =   25008
         Width           =   1095
      End
      Begin VB.CommandButton com 
         Caption         =   "&Ok"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   1
         Top             =   120
         WhatsThisHelpID =   25007
         Width           =   1095
      End
      Begin VB.CommandButton com 
         Caption         =   "Sfo&glia..."
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   0
         Top             =   1080
         WhatsThisHelpID =   25010
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSelUtente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine

'===============================================
'       definizione costanti
'===============================================
Const COM_OK = 0
Const COM_ANNULLA = 1
Const COM_SFOGLIA = 2

Enum enmStatoForm
    stsNormale = 0
    stsSfoglia = 1
End Enum

Enum enmTipoOperazione
    opeSeleziona = 0
    opeCopiaAccessi = 1
End Enum


'===============================================
'       definizione variabili
'===============================================
'variabili di stato della form
Dim MsetTipoOpe As enmTipoOperazione
Dim MsetStato As enmStatoForm
Dim mBolAnnulla As Boolean
Dim MbolIncludiSupervisor As Boolean

Sub resetcampi()
    txtb.text = ""
    txtb.tag = ""
    trwUtenti.Nodes.Clear
    com(COM_OK).Enabled = False
    setStato = stsNormale
    mBolAnnulla = False
End Sub

'NOME           : SelezionaGruppoUtente
'DESCRIZIONE    : seleziona un gruppo/utente
'PARAMETRO 1    : tipo operazione (vedi costante enumerativa)
'PARAMETRO 2    : (ritorno) gruppo/utente selezionato
'PARAMETRO 3    : (ritorno) G se selezionato un gruppo; U se selezionato un utente
'PARAMETRO 4    : true - include utenti supervisor
Public Function SelezionaGruppoUtente(ByVal setTipoOpe As enmTipoOperazione, _
                                        strKeySel As String, _
                                        strTipoRes As String, _
                                        strDscSel As String, _
                                        bolIncludiSup) As Boolean
    
    
    MbolIncludiSupervisor = bolIncludiSup
    MsetTipoOpe = setTipoOpe
    Me.Show vbModal
    
    SelezionaGruppoUtente = (Not mBolAnnulla)
    If (SelezionaGruppoUtente) Then
        strTipoRes = Left$(txtb.tag, 1)
        strKeySel = Mid$(txtb.tag, 2)
        strDscSel = txtb.text
    End If
    Unload Me
    
End Function

Property Let setStato(setNuovoStato As enmStatoForm)
    'impostazioni form
    If (setNuovoStato = stsNormale) Then
        Me.Height = 2025
        com(COM_SFOGLIA).Caption = MXNU.CaricaStringaRes(25010) & " >>"
    Else
        Me.Height = 5445
        com(COM_SFOGLIA).Caption = "<< " & MXNU.CaricaStringaRes(25021)
    End If
    Scheda.Height = Me.Height - 420
    'assegno la variabile
    MsetStato = setNuovoStato
End Property
Property Get setStato() As enmStatoForm
    setStato = MsetStato
End Property

Sub DefLingua()
    
    Me.Caption = MXNU.CaricaStringaRes(23007)
    Call MXNU.LeggiRisorseControlli(Me)
'    com(COM_OK).Caption = MXNU.CaricaStringaRes(25007)
'    com(COM_ANNULLA).Caption = MXNU.CaricaStringaRes(25008)
'    com(COM_SFOGLIA).Caption = MXNU.CaricaStringaRes(25010)
'    If (MsetTipoOpe = opeSeleziona) Then Frame(0).WhatsThisHelpID = 24026 Else Frame(0).WhatsThisHelpID = 24027
'    Frame(0).Caption = MXNU.CaricaStringaRes(Frame(0).WhatsThisHelpID)
'    Frame(1).Caption = MXNU.CaricaStringaRes(24023)
End Sub

Private Sub com_Click(Index As Integer)
    Select Case Index
        Case COM_OK
            mBolAnnulla = False
            Me.Hide
        Case COM_ANNULLA
            mBolAnnulla = True
            Me.Hide
        Case COM_SFOGLIA
            If (setStato = stsNormale) Then setStato = stsSfoglia Else setStato = stsNormale
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then Unload Me
End Sub


Private Sub Form_Load()
    
    Call DefLingua
    Call resetcampi
    
    Call TreeUtenti_Inizializza(trwUtenti, False, MbolIncludiSupervisor)
    trwUtenti.Nodes("gruppi").Expanded = False
    trwUtenti.Nodes("utenti").Expanded = False
    Call CentraFinestra(Me.hwnd)
    ' RIF.A#7261
    trwUtenti.Height = Frame(1).Height - 400
    
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSelUtente = Nothing
End Sub

Private Sub Scheda_Paint()
    Call SchedaOmbreggiaControlli(Scheda)
End Sub

Private Sub trwUtenti_DblClick()
    If (trwUtenti.SelectedItem <> trwUtenti.Nodes("Metodo98") And trwUtenti.SelectedItem <> trwUtenti.Nodes("gruppi") And trwUtenti.SelectedItem <> trwUtenti.Nodes("utenti")) Then
        txtb.text = trwUtenti.SelectedItem.text
        txtb.tag = trwUtenti.SelectedItem.tag
        Call txtb_LostFocus
    End If
End Sub


Private Sub txtb_KeyPress(KeyAscii As Integer)
    Call CtrlKey(KeyAscii, CKEY_CARASCII)
End Sub

Private Sub txtb_LostFocus()
    'validazione utente
    Dim strKey As String
    com(COM_OK).Enabled = ValidaUtenteGruppo(txtb.text, strKey)
    txtb.tag = strKey
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
'Per Metodo Evolus
Private Sub mResize_AfterResize()
    Call AvvicinaLing(Me)
End Sub

