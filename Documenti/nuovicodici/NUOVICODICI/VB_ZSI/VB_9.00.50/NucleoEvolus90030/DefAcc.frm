VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4AFA2505-EEFF-4BA2-873D-9FDF23CDB0CB}#1.0#0"; "MXCTRL.OCX"
Begin VB.Form frmDefAcc 
   Caption         =   "Definizione Accessi Gruppi/Utenti"
   ClientHeight    =   4935
   ClientLeft      =   1275
   ClientTop       =   2190
   ClientWidth     =   9255
   Icon            =   "DefAcc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9255
   Begin MXCTRL.MWSchedaBox Scheda 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8705
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
      ScaleWidth      =   9255
      ScaleHeight     =   4935
      Begin MXCTRL.MWSchedaBox pnlBack 
         Height          =   3255
         Left            =   6600
         TabIndex        =   4
         Top             =   600
         WhatsThisHelpID =   24025
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5741
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
         ScaleWidth      =   2415
         ScaleHeight     =   3255
         Caption         =   "Definizione Accessi"
         VAlign          =   0
         FillWithGradient=   0   'False
         Begin VB.CommandButton ComClr 
            Caption         =   "<Cancella Impostazioni>"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   2640
            WhatsThisHelpID =   25018
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Appearance      =   0  'Flat
            Caption         =   "&Annullamento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   10
            Top             =   2040
            WhatsThisHelpID =   50006
            Width           =   1695
         End
         Begin VB.CheckBox chk 
            Appearance      =   0  'Flat
            Caption         =   "&Inserimento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   9
            Top             =   1800
            WhatsThisHelpID =   50005
            Width           =   1695
         End
         Begin VB.CheckBox chk 
            Appearance      =   0  'Flat
            Caption         =   "&Modifica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   8
            Top             =   1560
            WhatsThisHelpID =   50004
            Width           =   1695
         End
         Begin VB.CheckBox chk 
            Appearance      =   0  'Flat
            Caption         =   "&Lettura"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   7
            Top             =   1320
            WhatsThisHelpID =   50003
            Width           =   1695
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Accesso &Disabilitato"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   480
            WhatsThisHelpID =   60003
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Accesso &Abilitato"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   960
            WhatsThisHelpID =   60004
            Width           =   2055
         End
      End
      Begin VB.CommandButton com 
         Caption         =   "<Copia su>"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   6960
         TabIndex        =   3
         Top             =   4440
         WhatsThisHelpID =   25020
         Width           =   1935
      End
      Begin VB.CommandButton com 
         Caption         =   "<Imposta>"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   4440
         WhatsThisHelpID =   25019
         Width           =   1935
      End
      Begin VB.CommandButton com 
         Caption         =   "<Annulla>"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   1
         Top             =   4440
         WhatsThisHelpID =   25008
         Width           =   1935
      End
      Begin VB.Frame frmFrame 
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
         Height          =   4095
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   120
         WhatsThisHelpID =   24023
         Width           =   3015
         Begin MSComctlLib.TreeView trwUtenti 
            Height          =   3735
            Left            =   120
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
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
      End
      Begin VB.Frame frmFrame 
         Appearance      =   0  'Flat
         Caption         =   "Nome Videata/SottoScheda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Index           =   1
         Left            =   3360
         TabIndex        =   13
         Top             =   120
         WhatsThisHelpID =   24024
         Width           =   3015
         Begin MSComctlLib.TreeView trwOggetti 
            Height          =   3735
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
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
      End
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
            Picture         =   "DefAcc.frx":0442
            Key             =   "metodo98"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":09E8
            Key             =   "entire"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":0AFA
            Key             =   "gruppo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":10A0
            Key             =   "utente"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":1646
            Key             =   "formab"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":1758
            Key             =   "formds"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":1CAA
            Key             =   "lingab"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":1DA4
            Key             =   "lingds"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":1E9E
            Key             =   "sitab"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":1FB8
            Key             =   "sitds"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":20D2
            Key             =   "findab"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAcc.frx":2224
            Key             =   "findds"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDefAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1

'===============================================
'       definizione costanti
'===============================================
Const COM_IMPOSTA = 0
Const COM_ANNULLA = 1
Const COM_COPIA = 2

Const OPT_DISABILITA = 0
Const OPT_ABILITA = 1

Const CHK_LETTURA = 0
Const CHK_MODIFICA = 1
Const CHK_INSERISCI = 2
Const CHK_ANNULLA = 3

Enum enmStato
    stsReadOnly = 0
    stsWrite = 1
End Enum
'===============================================
'       definizione variabili
'===============================================
Public frmDef As Form
Dim MlngIDForm As Long
Dim MstrKeyForm As String
Dim MstrCaption As String
Dim MstrSitEntry As String
Dim MbolAbilitaChk As Boolean
Dim mBolSit As Boolean

Dim setStato As enmStato

Dim colAccGruppi As Collection
Dim colAccUtenti As Collection

Private Sub AbilitaChkModifica()
    If opt(OPT_ABILITA).Value Then
        If chk(CHK_INSERISCI).Value = vbChecked Then
            chk(CHK_MODIFICA).Enabled = False
            chk(CHK_MODIFICA).Value = vbChecked
        Else
            chk(CHK_MODIFICA).Enabled = True
        End If
    End If
End Sub

'==============================================================================================
'                       eventi oggetti della form
'==============================================================================================
Private Sub Form_Initialize()
    Set colAccUtenti = New Collection
    Set colAccGruppi = New Collection
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then Unload Me
End Sub

Private Sub Form_Load()
    setStato = stsWrite
    
    MbolAbilitaChk = True
#If ISMETODO2005 = 1 Then
    If (frmDef.Name = "EmptyForm") Or (frmDef.Name = "frmMenu") Then
#Else
    If frmDef Is frmModuli Then
#End If
        MbolAbilitaChk = False
        If Not frmModuli.TrwModuli.SelectedItem Is Nothing Then
            If frmModuli.TrwModuli.SelectedItem.Child Is Nothing Or frmModuli.TrwModuli.SelectedItem Is frmModuli.TrwModuli.SelectedItem.Root Then
                Unload Me
                Exit Sub
            Else
                With frmModuli.TrwModuli.SelectedItem
                    MstrKeyForm = "M" & .key
                    MlngIDForm = .Tag
                    MstrCaption = .text
                    MstrSitEntry = ""
                End With
            End If
        Else
            Unload Me
            Exit Sub
        End If
    Else
        MlngIDForm = frmDef.HelpContextID
        MstrKeyForm = "F" & frmDef.Name
        MstrSitEntry = frmDef.Name
        MstrCaption = frmDef.Caption
    End If
    If (MlngIDForm <> 0) Then
        'imposto la lingua
        Call DefLingua
        'leggo e imposto gli accessi
        Call InizializzaStrutture
        'imposto gli alberi
        Call TreeUtenti_Inizializza(trwUtenti, True, False)
        Call TreeOggetti_Inizializza
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
        'mostro la finestra
        Call CentraFinestra(Me.hwnd)
    Else
        MsgBox MXNU.CaricaStringaRes(1038), vbCritical
        Unload Me
    End If
    If Not MXNU.CtrlAccessi Then
        Call MXNU.WritePrivacyLog(AvvioPermessiAppl, MXNU.CaricaStringaRes(3206))
    End If
End Sub

Private Sub Form_Terminate()
    Set colAccUtenti = Nothing
    Set colAccGruppi = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDefAcc = Nothing
End Sub

Private Sub Chk_Click(Index As Integer)
     
    If Index = CHK_INSERISCI Then   'Anomalia 6483
        Call AbilitaChkModifica
    End If
    
    If (setStato = stsReadOnly) Then Exit Sub
    
    
    Call Accessi_Write(trwUtenti.SelectedItem, trwOggetti.SelectedItem, True, (opt(OPT_ABILITA).Value), (chk(CHK_LETTURA).Value <> 0), (chk(CHK_MODIFICA).Value <> 0), (chk(CHK_INSERISCI).Value <> 0), (chk(CHK_ANNULLA).Value <> 0))
    
End Sub

Private Sub com_Click(Index As Integer)
    Select Case Index
        Case COM_IMPOSTA
            Call SalvaImpostazioni
            Unload Me
        Case COM_ANNULLA
            Unload Me
        Case COM_COPIA
           Call Accessi_Copy
    End Select
End Sub

Private Sub ComClr_Click()
    'Rif. anomalia #8240 (nodo non selezionato)
    If Not (trwUtenti.SelectedItem Is Nothing) Then
        'cancello le impostazioni del nodo selezionato
        Call Accessi_Write(trwUtenti.SelectedItem, trwOggetti.SelectedItem, False)
        Call MXNU.WritePrivacyLog(DisabilitaUtente, MXNU.CaricaStringaRes(3208, Array("", trwUtenti.SelectedItem, MstrCaption, -1)))
        'refresh dell'albero degli oggetti
        Call TreeOggetti_Refresh(trwUtenti.SelectedItem)
        Call trwOggetti_NodeClick(trwOggetti.SelectedItem)
    Else
        Call MXNU.MsgBoxEX(3068, vbExclamation, 1007)
    End If
End Sub

Private Sub opt_Click(Index As Integer)
Dim cnt As Integer
    
    For cnt = CHK_LETTURA To CHK_ANNULLA
        chk(cnt).Enabled = (Index = OPT_ABILITA) And MbolAbilitaChk
        If (Index = OPT_DISABILITA) Then chk(cnt).Value = 0
    Next cnt
    
    If (setStato = stsWrite) Then
        'leggo gli accessi precedentemente memorizzati
        Dim bolImpostato As Boolean, bolAccesso As Boolean, bolLettura As Boolean, bolModifica As Boolean, bolInserisci As Boolean, bolAnnulla As Boolean
        Call Accessi_Read(trwUtenti.SelectedItem, trwOggetti.SelectedItem, bolImpostato, bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)
        'abilito/disabilito nodo
        Call Nodo_Enable(trwOggetti.SelectedItem, opt(OPT_ABILITA).Value)
        'se impostazione precedente = non impostato o disabilitato -> abilito tutti gli accessi
        If opt(OPT_ABILITA).Value = True And (Not bolImpostato Or Not bolAccesso) Then
            setStato = stsReadOnly
            For cnt = CHK_LETTURA To CHK_ANNULLA
                If (chk(cnt).Visible) Then chk(cnt).Value = 1
            Next cnt
            setStato = stsWrite
            Call Accessi_Write(trwUtenti.SelectedItem, trwOggetti.SelectedItem, True, (opt(OPT_ABILITA).Value), (chk(CHK_LETTURA).Value <> 0), (chk(CHK_MODIFICA).Value <> 0), (chk(CHK_INSERISCI).Value <> 0), (chk(CHK_ANNULLA).Value <> 0))
        End If
    End If
    
End Sub

Private Sub Scheda_Paint()
    Call SchedaOmbreggiaControlli(Scheda())
End Sub

Private Sub trwOggetti_NodeClick(ByVal Node As MSComctlLib.Node)
Dim bolImpostato As Boolean, bolAccesso As Boolean, bolLettura As Boolean, bolModifica As Boolean, bolInserisci As Boolean, bolAnnulla As Boolean
    
    'aggiorno il pannello a seconda del nodo selezionato
    Call pnlBack_Refresh(Node)
    'leggo gli accessi per il nodo selezionato
    If (Accessi_Read(trwUtenti.SelectedItem, Node, bolImpostato, bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)) Then
        Call NodoPaint_Imposta(Node, bolImpostato, bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)
    End If
    'nodo form disabilitato -> disabilito il pannello
    Select Case Node.key
        Case MstrKeyForm
            Call pnlBack_Enable(True)
        Case "lingsit"
            Call pnlBack_Enable(False)
        Case Else
            Call pnlBack_Enable(True)
            If (Accessi_Read(trwUtenti.SelectedItem, Node.Parent, bolImpostato, bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)) Then
                If bolImpostato Then Call pnlBack_Enable(bolAccesso)
            End If
    End Select
End Sub

Private Sub trwUtenti_NodeClick(ByVal Node As MSComctlLib.Node)
    'refresh dell'albero degli oggetti
    Call TreeOggetti_Refresh(Node)
    'leggo gli accessi per il nodo selezionato
    Dim bolImpostato As Boolean, bolAccesso As Boolean, bolLettura As Boolean, bolModifica As Boolean, bolInserisci As Boolean, bolAnnulla As Boolean
    If (Accessi_Read(Node, trwOggetti.SelectedItem, bolImpostato, bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)) Then
        'abilito il bottone copia
        com(COM_COPIA).Enabled = True
        'imposto nodo
        Call NodoPaint_Imposta(trwOggetti.SelectedItem, bolImpostato, bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)
    Else
        'disabilito il bottone copia
        com(COM_COPIA).Enabled = False
        'imposto nodo
        trwOggetti.SelectedItem = trwOggetti.Nodes(MstrKeyForm)
        Call trwOggetti_NodeClick(trwOggetti.Nodes(MstrKeyForm))
    End If
    
End Sub


'==============================================================================================
'                       funzioni private della form
'==============================================================================================
Private Sub DefLingua()

    Me.Caption = MXNU.CaricaStringaRes(23006)
'    frmFrame(0).Caption = MXNU.CaricaStringaRes(24023)
'    frmFrame(1).Caption = MXNU.CaricaStringaRes(24024)
'    pnlBack.Caption = MXNU.CaricaStringaRes(24025)
'    opt(OPT_DISABILITA).Caption = MXNU.CaricaStringaRes(60003)
'    opt(OPT_ABILITA).Caption = MXNU.CaricaStringaRes(60004)
'    chk(CHK_LETTURA).Caption = MXNU.CaricaStringaRes(50003)
'    chk(CHK_MODIFICA).Caption = MXNU.CaricaStringaRes(50004)
'    chk(CHK_INSERISCI).Caption = MXNU.CaricaStringaRes(50005)
'    chk(CHK_ANNULLA).Caption = MXNU.CaricaStringaRes(50006)
'    ComClr.Caption = MXNU.CaricaStringaRes(25018)
'    com(COM_IMPOSTA).Caption = MXNU.CaricaStringaRes(25019)
'    com(COM_ANNULLA).Caption = MXNU.CaricaStringaRes(25008)
'    com(COM_COPIA).Caption = MXNU.CaricaStringaRes(25020)
    Call MXNU.LeggiRisorseControlli(Me)
End Sub

Private Sub InizializzaStrutture()
Dim intq As Integer
Dim strSQL As String
Dim hSS As CRecordSet
Dim bolEnd As Boolean
Dim vntAus As Variant
Dim CAcc As CAccUtenti
Dim lngIDForm As Long
Dim vetSchede() As Integer
Dim vetVisSit() As Integer

    'creo l'array con gli indici da inizializzare
    Call DammiOggetti(lngIDForm, vetSchede(), vetVisSit())
    'inizializzo struttua gruppi
    strSQL = "SELECT Codice" _
            & " FROM TabGruppiUtente"
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    Do While (Not bolEnd)
        vntAus = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Codice", "")
        If (vntAus <> "") Then
            Set CAcc = New CAccUtenti
            Call CAcc.Inizializza(vntAus, tipGruppo, lngIDForm, vetSchede(), mBolSit, vetVisSit(), MstrSitEntry, MstrCaption)
            colAccGruppi.Add CAcc, "G" & vntAus
        End If
        bolEnd = Not MXDB.dbSuccessivo(hSS, TIPO_SNAPSHOT)
    Loop
    intq = MXDB.dbChiudiSS(hSS)
    
    'inizializzo struttua utenti
    strSQL = "SELECT UserID" _
            & " FROM TabUtenti" _
            & " WHERE Supervisor = 0"
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    Do While (Not bolEnd)
        vntAus = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "UserID", "")
        If (vntAus <> "") Then
            Set CAcc = New CAccUtenti
            Call CAcc.Inizializza(vntAus, tipUtente, lngIDForm, vetSchede(), mBolSit, vetVisSit(), MstrSitEntry, MstrCaption)
            colAccUtenti.Add CAcc, "U" & vntAus
        End If
        bolEnd = Not MXDB.dbSuccessivo(hSS, TIPO_SNAPSHOT)
    Loop
    intq = MXDB.dbChiudiSS(hSS)
End Sub

Private Sub DammiOggetti(lngIDForm As Long, _
                        vetSchede() As Integer, _
                        vetVisSit() As Integer)
Dim ctrGen As Control
Dim intSch As Integer
Dim strAus As String

    'identificativo form
    lngIDForm = MlngIDForm 'frmDef.HelpContextID
    'indici schede
    intSch = -1
    ReDim vetSchede(0) As Integer
    For Each ctrGen In ControlliForm(frmDef)
        If (TypeName(ctrGen) = "MWLinguetta") Then
            If (StrComp(ctrGen.Name, "Ling", vbTextCompare) = 0) Or (StrComp(ctrGen.Name, "LingIva", vbTextCompare) = 0) Then
                intSch = intSch + 1
                ReDim Preserve vetSchede(intSch) As Integer
                If StrComp(ctrGen.Name, "Ling", vbTextCompare) = 0 Then
                    vetSchede(intSch) = ctrGen.Index
                Else
                    'Imposto l'indice scheda a 999+indice linguetta per le linguette iva (LingIva) della prima nota
                    'Anomalia nr. 3639: vedi funzione TreeOggetti_LingGet
                    vetSchede(intSch) = 9990 + ctrGen.Index
                End If
            End If
        End If
    Next
    'scheda situazione
    mBolSit = False
    ReDim vetVisSit(0) As Integer
    For Each ctrGen In ControlliForm(frmDef)
        If (TypeName(ctrGen) = "MWLinguetta") Then
            If (StrComp(ctrGen.Name, "LingSit", vbTextCompare) = 0) Then
                vetVisSit(0) = 0
                Dim cSit As MXKit.cSituazione
                Dim intSit As Integer
                'leggo le visioni legate alla form
                Set cSit = MXVI.CreaCSituazione()
                cSit.pCtrlAccessi = False
                mBolSit = cSit.LeggiSituazioniDisponibili(LeggiNomeSituazione)
                If mBolSit Then
                    ReDim Preserve vetVisSit(cSit.pSituazioniDisponibili.Count) As Integer
                    For intSit = 1 To cSit.pSituazioniDisponibili.Count
                        vetVisSit(intSit) = -(intSit)
                    Next intSit
                End If
                Set cSit = Nothing
            End If
        End If
    Next
End Sub

Private Function LeggiNomeSituazione() As String
    Dim strNome As String
    
    Select Case frmDef.Name
        Case "OrdProd"
            strNome = "SITORDPROD"
        Case "frmSituazione"
            'RIF. A#6560 - Gestisce la finestra generica delle situazioni
            strNome = frmDef.pNomeSituazione
        Case Else
            strNome = frmDef.Name
    End Select
    
    LeggiNomeSituazione = strNome
End Function

Private Sub Accessi_Copy()
Dim vetParam() As Variant
Dim strSrcKey As String, strSrcTip As String, strSrcDsc As String
Dim strDstKey As String, strDstTip As String, strDstDsc As String
Dim CAccSrc As CAccUtenti, CAccDst As CAccUtenti

    strSrcTip = Left$(trwUtenti.SelectedItem.Tag, 1)
    strSrcKey = Mid$(trwUtenti.SelectedItem.Tag, 2)
    strSrcDsc = trwUtenti.SelectedItem.text
    If (frmSelUtente.SelezionaGruppoUtente(opeCopiaAccessi, strDstKey, strDstTip, strDstDsc, False)) Then
        ReDim vetParam(1 To 4) As Variant
        If (strSrcTip = "G") Then vetParam(1) = MXNU.CaricaStringaRes(24029) Else vetParam(1) = MXNU.CaricaStringaRes(24030)
        vetParam(2) = strSrcDsc
        If (strDstTip = "G") Then vetParam(3) = MXNU.CaricaStringaRes(24029) Else vetParam(3) = MXNU.CaricaStringaRes(24030)
        vetParam(4) = strDstDsc
        If (MsgBox(MXNU.CaricaStringaRes(1036, vetParam()), vbQuestion + vbYesNo) = vbYes) Then
            If (strSrcTip = "G") Then Set CAccSrc = colAccGruppi(strSrcTip & strSrcKey) Else Set CAccSrc = colAccUtenti(strSrcTip & strSrcKey)
            If (strDstTip = "G") Then Set CAccDst = colAccGruppi(strDstTip & strDstKey) Else Set CAccDst = colAccUtenti(strDstTip & strDstKey)
            Call CAccDst.CopiaAccessi(CAccSrc)
            Set CAccSrc = Nothing
            Set CAccDst = Nothing
        End If
    End If
    
End Sub

Private Function Accessi_Read(ByVal nodUtente As MSComctlLib.Node, _
                        ByVal nodOggetto As MSComctlLib.Node, _
                        bolImpostato As Boolean, _
                        bolAccesso As Boolean, _
                        bolLettura As Boolean, _
                        bolModifica As Boolean, _
                        bolInserisci As Boolean, _
                        bolAnnulla As Boolean) As Boolean
                        
    If (nodUtente Is Nothing Or nodOggetto Is Nothing) Then
        Accessi_Read = False
    Else
        If (nodUtente = trwUtenti.Nodes("Metodo98") Or nodUtente = trwUtenti.Nodes("gruppi") Or nodUtente = trwUtenti.Nodes("utenti") Or nodOggetto.key = "lingsit") Then
            Accessi_Read = False
        Else
            'accessi utente/gruppo
            Dim colAccessi As Collection
            Accessi_Read = True
            If (Left$(nodUtente.Tag, 1) = "G") Then
                Set colAccessi = colAccGruppi
            Else
                Set colAccessi = colAccUtenti
            End If
            If (Left$(nodOggetto.key, 1) = "S") Then
                'visione situazione
                bolImpostato = (colAccessi(CStr(nodUtente.Tag)).AccessiImpostati(ID_SCHEDA_SITUAZIONE, Val(nodOggetto.Tag))) '(rif.sch.2302)
                If bolImpostato Then
                    Call colAccessi(CStr(nodUtente.Tag)).LeggiAccessi(ID_SCHEDA_SITUAZIONE, Val(nodOggetto.Tag), bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)
                End If
            Else
                'linguetta
                bolImpostato = (colAccessi(CStr(nodUtente.Tag)).AccessiImpostati(MlngIDForm, Val(nodOggetto.Tag)))
                If bolImpostato Then
                    Call colAccessi(CStr(nodUtente.Tag)).LeggiAccessi(MlngIDForm, Val(nodOggetto.Tag), bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)
                End If
            End If
            Set colAccessi = Nothing
        End If
    End If
    
End Function

Private Sub Accessi_Write(ByVal nodUtente As MSComctlLib.Node, _
                        ByVal nodOggetto As MSComctlLib.Node, _
                        bolImpostato As Boolean, _
                        Optional bolAccesso As Boolean, _
                        Optional bolLettura As Boolean, _
                        Optional bolModifica As Boolean, _
                        Optional bolInserisci As Boolean, _
                        Optional bolAnnulla As Boolean)
    
    If Not (nodUtente Is Nothing Or nodOggetto Is Nothing) Then
        If (nodUtente <> trwUtenti.Nodes("Metodo98") And nodUtente <> trwUtenti.Nodes("gruppi") And nodUtente <> trwUtenti.Nodes("utenti")) Then
            Dim colAccessi As Collection
            If (Left$(nodUtente.Tag, 1) = "G") Then
                Set colAccessi = colAccGruppi
            Else
                Set colAccessi = colAccUtenti
            End If
            If (Left$(nodOggetto.key, 1) = "S") Then
                'visione situazione
                Call colAccessi(CStr(nodUtente.Tag)).MemorizzaAccessi(ID_SCHEDA_SITUAZIONE, Val(nodOggetto.Tag), bolImpostato, bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)
            Else
                'linguetta
                Call colAccessi(CStr(nodUtente.Tag)).MemorizzaAccessi(MlngIDForm, Val(nodOggetto.Tag), bolImpostato, bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)
            End If
        End If
    End If
End Sub

Private Sub Nodo_Enable(ByVal nodOggetto As MSComctlLib.Node, ByVal bolEnable As Boolean)
Dim intNode As Integer

    If nodOggetto Is Nothing Then Exit Sub
    Call NodoPaint_Enable(nodOggetto, bolEnable, False)
    If (nodOggetto = trwOggetti.SelectedItem) Then
        'oggetto selezionato -> scrivo abilitazioni
        Call Accessi_Write(trwUtenti.SelectedItem, nodOggetto, True, (opt(OPT_ABILITA).Value), (chk(CHK_LETTURA).Value <> 0), (chk(CHK_MODIFICA).Value <> 0), (chk(CHK_INSERISCI).Value <> 0), (chk(CHK_ANNULLA).Value <> 0))
    Else
        'oggetto figlio di quello selezionato
        If (bolEnable) Then
            'nodo padre abilitato -> cancello le impostazioni
            Call Accessi_Write(trwUtenti.SelectedItem, nodOggetto, False)
        Else
            'nodo padre disabilitato -> disabilito il nodo figlio
            Call Accessi_Write(trwUtenti.SelectedItem, nodOggetto, True, False)
        End If
    End If
    'processo i figli
    If (nodOggetto.children > 0) Then
        intNode = nodOggetto.Child.Index
        While (intNode <> nodOggetto.Child.LastSibling.Index)
            Call Nodo_Enable(trwOggetti.Nodes(intNode), bolEnable)
            'prossimo nodo
            intNode = trwOggetti.Nodes(intNode).Next.Index
        Wend
        intNode = nodOggetto.Child.LastSibling.Index
        Call Nodo_Enable(trwOggetti.Nodes(intNode), bolEnable)
    End If

End Sub

Private Sub NodoPaint_Enable(ByVal Node As MSComctlLib.Node, bolEnable As Boolean, bolProcChild As Boolean)
    If (Node Is Nothing) Then Exit Sub
    
    Dim intNode As Integer
    Dim strImage As String
    'modifico l'immagine
    strImage = Left$(Node.Image, Len(Node.Image) - 2)
    If (bolEnable) Then Node.Image = strImage & "ab" Else Node.Image = strImage & "ds"
    'processo i figli
    If (bolProcChild And Node.children > 0) Then
        intNode = Node.Child.Index
        While (intNode <> Node.Child.LastSibling.Index)
            Call NodoPaint_Enable(trwOggetti.Nodes(intNode), bolEnable, True)
            'prossimo nodo
            intNode = trwOggetti.Nodes(intNode).Next.Index
        Wend
        intNode = Node.Child.LastSibling.Index
        Call NodoPaint_Enable(trwOggetti.Nodes(intNode), bolEnable, True)
    End If
End Sub

Private Sub NodoPaint_Imposta(ByVal Node As MSComctlLib.Node, _
                                bolImpostato As Boolean, _
                                bolAccesso As Boolean, _
                                bolLettura As Boolean, _
                                bolModifica As Boolean, _
                                bolInserisci As Boolean, _
                                bolAnnulla As Boolean)
                                    
    setStato = stsReadOnly
    If (bolImpostato) Then
        If (bolAccesso) Then
            opt(OPT_ABILITA).Value = True
            chk(CHK_LETTURA).Enabled = MbolAbilitaChk
            chk(CHK_MODIFICA).Enabled = MbolAbilitaChk
            chk(CHK_INSERISCI).Enabled = MbolAbilitaChk
            chk(CHK_ANNULLA).Enabled = MbolAbilitaChk
        Else
            opt(OPT_DISABILITA).Value = True
            chk(CHK_LETTURA).Enabled = False
            chk(CHK_MODIFICA).Enabled = False
            chk(CHK_INSERISCI).Enabled = False
            chk(CHK_ANNULLA).Enabled = False
        End If
        Call NodoPaint_Enable(Node, bolAccesso, False)
        chk(CHK_LETTURA).Value = Abs(bolLettura)
        chk(CHK_MODIFICA).Value = Abs(bolModifica)
        chk(CHK_INSERISCI).Value = Abs(bolInserisci)
        Call AbilitaChkModifica   'Anomalia 6483
        chk(CHK_ANNULLA).Value = Abs(bolAnnulla)
    Else
        opt(OPT_ABILITA).Value = False
        opt(OPT_DISABILITA).Value = False
        chk(CHK_LETTURA).Value = False
        chk(CHK_MODIFICA).Value = False
        chk(CHK_INSERISCI).Value = False
        chk(CHK_ANNULLA).Value = False
    End If
    setStato = stsWrite
                                
End Sub

Private Sub pnlBack_Enable(bolEnabled As Boolean)
Dim cnt As Integer
    'disabilito il pannello...
    pnlBack.Enabled = bolEnabled
    '... e tutti i controlli in esso contenuti
    For cnt = 0 To 1
        opt(cnt).Enabled = bolEnabled
    Next cnt
    If (opt(OPT_ABILITA).Value) Then
        For cnt = 0 To 3
            chk(cnt).Enabled = bolEnabled And MbolAbilitaChk
        Next cnt
        If bolEnabled And MbolAbilitaChk Then
            If chk(CHK_INSERISCI).Visible Then   'Anomalia 6483
                Call AbilitaChkModifica
            End If
        End If
    End If
    DoEvents
End Sub

Private Sub pnlBack_Refresh(ByVal Node As MSComctlLib.Node)
Dim cnt As Integer
    setStato = stsReadOnly
    For cnt = CHK_LETTURA To CHK_ANNULLA
        chk(cnt).Visible = True
    Next cnt
    For cnt = OPT_DISABILITA To OPT_ABILITA
        opt(cnt).Visible = True
    Next cnt
    
    If (Node = trwOggetti.Nodes(MstrKeyForm)) Then
        'selezionata la form -> visualizzo i check inserisci/annulla
    ElseIf (Left$(Node.key, 1) = "S") Then
        'selezionata visione -> visualizzo solo abilita/disabilita
        For cnt = CHK_MODIFICA To CHK_ANNULLA
            chk(cnt).Visible = False
            chk(cnt).Value = 0
        Next cnt
    Else
        'selezionata una linguetta -> nascondo i check inserisci/annulla
        chk(CHK_INSERISCI).Visible = False
        chk(CHK_ANNULLA).Visible = False
        chk(CHK_INSERISCI).Value = 0
        chk(CHK_ANNULLA).Value = 0
    End If
    setStato = stsWrite
End Sub

Private Sub TreeOggetti_LingGet(ctrParent As Object, ByVal strKeyParent As String)
    Dim nodX As Node
    Dim ctrGen As Control
    Dim strLingKey As String
    Dim schede As Object
    Dim bolValido As Boolean
    Static sbolPassato As Boolean

    On Local Error Resume Next
    Set schede = ContenitoreControlli(frmDef).Controls("Scheda")
    On Local Error GoTo 0
    
    '**************************************************************************************
    'ATTENZIONE: Gestione caso particolare per form Prima Nota Contabile che contiene
    '            delle sottolinguette chiamate LingIva
    'Rif. Anomalie98 Nr. 3639
    '**************************************************************************************
    
    For Each ctrGen In ctrParent.Controls
        If (TypeName(ctrGen) = "MWLinguetta") Then
            'If (StrComp(ctrGen.Name, "Ling", vbTextCompare) = 0) And (ctrGen.Container Is ctrParent) And (ctrGen.Visible) Then
            bolValido = (StrComp(ctrGen.Name, "Ling", vbTextCompare) = 0) And (ctrGen.Container Is ctrParent)
            If Not bolValido Then
                'Rif. anomalia #8710 - alcune form con scheda come base non considerano il caricamento delle linguette
                Select Case LCase(frmDef.Name)
                    Case "frmpncontabile"
                        bolValido = (LCase(ctrGen.Container.Name) = "schtesta")
                        If Not bolValido Then
                            If LCase(ctrGen.Container.Name) = "scheda" Then
                                bolValido = (ctrGen.Container.Index = 1)
                            End If
                        End If
                    Case "frmvarianti"
                        bolValido = (LCase(ctrGen.Container.Name) = "scheda") And (ctrGen.Container.Index = 3)
                 End Select
            End If
            If bolValido Then
                strLingKey = "L" & ctrGen.Name & "_" & ctrGen.Index
                On Local Error Resume Next
                Set nodX = trwOggetti.Nodes.Add(strKeyParent, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
                If Err.Number = 0 Then
                    If ctrGen.Name <> "LingIva" Then
                        nodX.Tag = ctrGen.Index
                    Else
                        nodX.Tag = "999" & ctrGen.Index
                    End If
                    Call nodX.EnsureVisible
                    'cerca eventuali sottolinguette
                    If LCase(frmDef.Name) <> "frmpncontabile" Then
                        Call TreeOggetti_LingGet(schede(ctrGen.Index), strLingKey)
                    ElseIf Not sbolPassato And LCase(strLingKey) = "lling_1" Then
                        sbolPassato = True
                        Call TreeOggetti_LingGet(schede(ctrGen.Index), strLingKey)
                    End If
                End If
                Err.Clear
                On Local Error GoTo 0
            End If
        End If
    Next

End Sub

Private Sub TreeOggetti_Inizializza()
    Dim nodX As Node
    Dim ctrGen As Control
    Dim ctrParent As Object
    
    'imposto radice=form
    Set nodX = trwOggetti.Nodes.Add(, tvwFirst, MstrKeyForm, MstrCaption, "formab")
    nodX.Tag = CStr(ID_SCHEDA_FORM)
    Call nodX.EnsureVisible
    'imposto linguette
    Set ctrParent = ContenitoreControlli(frmDef)
    Call TreeOggetti_LingGet(ctrParent, MstrKeyForm)
    'cerco linguetta situazione
    If mBolSit Then Call TreeOggetti_LingSit
End Sub

Private Sub TreeOggetti_LingSit()
Dim ctrGen As Control
Dim nodX As Node
Dim cSit As MXKit.cSituazione
Dim intSit As Integer
Dim strDsc As String
    On Local Error GoTo fine_LingSit
    For Each ctrGen In ControlliForm(frmDef)
        If (TypeName(ctrGen) = "MWLinguetta") Then
            If (StrComp(ctrGen.Name, "LingSit", vbTextCompare) = 0) Then
                Set nodX = trwOggetti.Nodes.Add(MstrKeyForm, tvwChild, "S" & ctrGen.Name, swapp(ctrGen.Caption, "&", ""), "sitab")
                nodX.Tag = 0
                Call nodX.EnsureVisible
                'carica visioni situazione
                Set cSit = MXVI.CreaCSituazione()
                cSit.pCtrlAccessi = False
                If cSit.LeggiSituazioniDisponibili(LeggiNomeSituazione) Then
                    For intSit = 1 To cSit.pSituazioniDisponibili.Count
                        strDsc = cSit.pDatiSituazione(intSit).strCaption
                        Set nodX = trwOggetti.Nodes.Add("S" & ctrGen.Name, tvwChild, CStr("S" & intSit), strDsc, "findab")
                        nodX.Tag = -(intSit)
                    Next intSit
                End If
                Set cSit = Nothing
            End If
        End If
    Next
fine_LingSit:
    On Local Error GoTo 0
    Exit Sub
End Sub

'NOME           : TreeOggetti_Refresh
'DESCRIZIONE    : effettua il refresh dei nodi oggetti in base al nodo utente selezionato
Private Sub TreeOggetti_Refresh(ByVal Node As MSComctlLib.Node)
    If (Node Is Nothing) Then Exit Sub
    
    Dim nodX As Node
    Dim bolImpostato As Boolean, bolAccesso As Boolean, bolLettura As Boolean, bolModifica As Boolean, bolInserisci As Boolean, bolAnnulla As Boolean
    If (Node.key = "Metodo98" Or Node.key = "gruppi" Or Node.key = "utenti") Then
        trwOggetti.Enabled = False
        Call pnlBack_Enable(False)
        For Each nodX In trwOggetti.Nodes
            Call NodoPaint_Enable(nodX, False, False)
        Next nodX
    Else
        trwOggetti.Enabled = True
        Call pnlBack_Enable(True)
        For Each nodX In trwOggetti.Nodes
            If (Accessi_Read(trwUtenti.SelectedItem, nodX, bolImpostato, bolAccesso, bolLettura, bolModifica, bolInserisci, bolAnnulla)) Then
                If (bolImpostato) Then
                    Call NodoPaint_Enable(nodX, bolAccesso, False)
                Else
                    Call NodoPaint_Enable(nodX, True, False)
                End If
            End If
        Next nodX
    End If
End Sub


Private Sub SalvaImpostazioni()
Dim CAcc As CAccUtenti

    Screen.MousePointer = vbHourglass
    'salvo le impostazioni per i gruppi
    For Each CAcc In colAccGruppi
        Call CAcc.SalvaAccessiGruppo(MlngIDForm)
    Next
    'salvo le impostazioni per gli utenti
    For Each CAcc In colAccUtenti
        Call CAcc.SalvaAccessiUtente(MlngIDForm, MstrSitEntry)
    Next
    Screen.MousePointer = vbDefault
    
End Sub
'==============================================================================================
'                       funzioni pubbliche della form
'==============================================================================================


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

