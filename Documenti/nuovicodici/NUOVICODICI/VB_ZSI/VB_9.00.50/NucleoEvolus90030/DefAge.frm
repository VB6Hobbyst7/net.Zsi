VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D4659C3-7027-4D03-BE01-53D83DB1A514}#1.0#0"; "MXCTRL.OCX"
Begin VB.Form frmDefAge 
   Caption         =   "Definizione Agenti Predefiniti"
   ClientHeight    =   4995
   ClientLeft      =   690
   ClientTop       =   2070
   ClientWidth     =   7335
   Icon            =   "DefAge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7335
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   4995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8811
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleWidth      =   7335
      ScaleHeight     =   4995
      Begin MXCtrl.MWSchedaBox pnlBack 
         Height          =   2895
         Left            =   4560
         TabIndex        =   4
         Top             =   1020
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5106
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScaleWidth      =   2415
         ScaleHeight     =   2895
         Caption         =   "Impostazione Agente"
         VAlign          =   0
         FillWithGradient=   0   'False
         Begin VB.CommandButton ComClr 
            Caption         =   "&Cancella Impostazioni"
            Height          =   375
            Left            =   300
            TabIndex        =   11
            Top             =   2280
            WhatsThisHelpID =   25018
            Width           =   1815
         End
         Begin VB.ComboBox cmbAgenti 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   420
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "&Esegui Agente Gruppo"
            Enabled         =   0   'False
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
            Left            =   180
            TabIndex        =   9
            Top             =   1440
            WhatsThisHelpID =   60010
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "&Ometti Agente Gruppo"
            Enabled         =   0   'False
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
            Left            =   180
            TabIndex        =   8
            Top             =   960
            WhatsThisHelpID =   60009
            Width           =   2055
         End
         Begin VB.PictureBox pnlP 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   300
            ScaleHeight     =   495
            ScaleWidth      =   1995
            TabIndex        =   5
            Top             =   1680
            Width           =   1995
            Begin VB.OptionButton optP 
               Appearance      =   0  'Flat
               Caption         =   "&Priorità &Utente"
               Enabled         =   0   'False
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
               Left            =   60
               TabIndex        =   7
               Top             =   240
               WhatsThisHelpID =   60012
               Width           =   1575
            End
            Begin VB.OptionButton optP 
               Appearance      =   0  'Flat
               Caption         =   "Priorità &Gruppo"
               Enabled         =   0   'False
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
               Left            =   60
               TabIndex        =   6
               Top             =   0
               WhatsThisHelpID =   60011
               Width           =   1575
            End
         End
      End
      Begin VB.CommandButton com 
         Caption         =   "<Annulla>"
         Height          =   375
         Index           =   1
         Left            =   2340
         TabIndex        =   3
         Top             =   4440
         WhatsThisHelpID =   25008
         Width           =   1935
      End
      Begin VB.CommandButton com 
         Caption         =   "<Imposta>"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   4440
         WhatsThisHelpID =   25019
         Width           =   1935
      End
      Begin VB.CommandButton com 
         Caption         =   "<Copia su>"
         Height          =   375
         Index           =   2
         Left            =   4740
         TabIndex        =   1
         Top             =   4440
         WhatsThisHelpID =   25020
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
         Top             =   180
         WhatsThisHelpID =   24023
         Width           =   4035
         Begin MSComctlLib.TreeView trwUtenti 
            Height          =   3735
            Left            =   120
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   3795
            _ExtentX        =   6694
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
      Begin MXCtrl.MWEtichetta etcfinestra 
         Height          =   375
         Left            =   4560
         Top             =   420
         Width           =   2415
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
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":0442
            Key             =   "metodo98"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":09E8
            Key             =   "entire"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":0AFA
            Key             =   "gruppo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":10A0
            Key             =   "utente"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":1646
            Key             =   "ctrl"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":1828
            Key             =   "evento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":1962
            Key             =   "form"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":1A74
            Key             =   "ling"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":1B86
            Key             =   "textbox"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":1D68
            Key             =   "maskedit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":1F4A
            Key             =   "button"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":212C
            Key             =   "checkbox"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":230E
            Key             =   "option"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":24F0
            Key             =   "combobox"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":26D2
            Key             =   "spreadsheet"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefAge.frx":28B4
            Key             =   "generico"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDefAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
DefLng A-Z

'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine

'===============================================
'       definizione costanti
'===============================================
'gestione alberi
'Const KEY_LING = "CTRL_Linguetta"
'Const KEY_TEXTBOX = "CTRL_TextBox"
'Const KEY_MASKEDIT = "CTRL_MaskEdit"
'Const KEY_BUTTON = "CTRL_CommandButton"
'Const KEY_CHECK = "CTRL_CheckBox"
'Const KEY_OPTION = "CTRL_OptionButton"
'Const KEY_COMBO = "CTRL_ComboBox"
'Const KEY_SPREAD = "CTRL_SpreadSheet"
'Const KEY_ALTRE = "CTRL_Altri"
'
'Const PIC_FORM = "form"
'Const PIC_CTRL = "metodo98"
'Const PIC_LING = "ling"
'Const PIC_TEXTBOX = "textbox"
'Const PIC_MASKEDIT = "maskedit"
'Const PIC_BUTTON = "button"
'Const PIC_CHECK = "checkbox"
'Const PIC_OPTION = "option"
'Const PIC_COMBO = "combobox"
'Const PIC_SPREAD = "spreadsheet"
'Const PIC_ALTRE = "ctrl"
'
'Enum enmCtrlTipoVisione
'    visPerTipo = 0
'    visPerScheda = 1
'End Enum

'controlli form
Const COM_IMPOSTA = 0
Const COM_ANNULLA = 1
Const COM_COPIA = 2

Const OPT_OMETTI = 0
Const OPT_ESEGUI = 1
Const OPT_ESE_GRUPPO = 0
Const OPT_ESE_UTENTE = 1

Enum enmAgtStato
    stsReadOnly = 0
    stsWrite = 1
End Enum
'===============================================
'       definizione variabili
'===============================================
Public frmDef As Form
Dim MlngIDForm As Long
Dim setStato As enmAgtStato

Dim colAgtGruppi As Collection
Dim colAgtUtenti As Collection

Dim bolOggImp As Boolean
Dim colOggetti As Collection

Sub Agenti_CaricaCombo()
    Dim strFile As String
    Dim strElenco As String
    Dim vntElenco As Variant
    Dim i%

    cmbAgenti.Clear
    'Call cmbAgenti.addItem("")
    On Local Error Resume Next
    strFile = Dir$(MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\Agenti\*.cmp", vbNormal)
    Do While (strFile <> "")
        If Not InStr(1, strElenco, strFile) Then strElenco = strElenco & "|" & strFile
        strFile = Dir$()
    Loop
    strFile = Dir$(MXNU.PercorsoPers & "\Agenti\*.cmp", vbNormal)
    Do While (strFile <> "")
        If Not InStr(1, strElenco, strFile) Then strElenco = strElenco & "|" & strFile
        strFile = Dir$()
    Loop
    strFile = Dir$(MXAA.PathAgenti & "\*.cmp", vbNormal)
    Do While (strFile <> "")
        If Not InStr(1, strElenco, strFile) Then strElenco = strElenco & "|" & strFile
        strFile = Dir$()
    Loop
    vntElenco = Split(strElenco, "|")
    For i = 0 To UBound(vntElenco)
        Call cmbAgenti.addItem(vntElenco(i))
    Next
    On Local Error GoTo 0

End Sub

'Sub AssegnaTag(nodX As Node, strNomeCtrl As String, setEvento As enmAgtEventi)
'
'    Dim CImp As CImposta
'
'    'assegna il tag al nodo
'    nodX.Tag = KeyAgentiGet(strNomeCtrl, setEvento)
'    If (Not bolOggImp) Then
'        'aggiunge l'elemento inserito alla collection
'        Set CImp = New CImposta
'        CImp.strNomeCtrl = strNomeCtrl
'        CImp.setEvento = setEvento
'        Call colOggetti.Add(CImp, nodX.Tag)
'    End If
'End Sub

Private Function CercaAgente(strAgente As String) As Integer

    Dim intLI As Integer

    CercaAgente = -1
    If strAgente <> "" Then
        For intLI = 0 To cmbAgenti.ListCount
            If (StrComp(strAgente, cmbAgenti.List(intLI), vbTextCompare) = 0) Then
                CercaAgente = intLI
                Exit For
            End If
        Next intLI
    End If
End Function

'Function CtrlGetDsc(ctrlGen As Control) As String
'
'   Dim vntIndex As Variant
'    Dim strNom As String
'    Dim strDsc As String
'
'    On Local Error Resume Next
'    vntIndex = ctrlGen.Index
'    'nome controllo
'    If (Err = 0) Then
'        strNom = ctrlGen.Name & "(" & vntIndex & ")"
'    Else
'        strNom = ctrlGen.Name
'    End If
'    'descrizione
'    Select Case TypeName(ctrlGen)
'        Case "MWLinguetta", "CommandButton", "CheckBox", "OptionButton"
'            strDsc = swapp(ctrlGen.Caption, "&", "")
'        Case "TextBox", "MaskEdBox"
'            strDsc = ctrlGen.Text
'        Case "ComboBox"
'            strDsc = ctrlGen.Text
'        Case Else
'            strDsc = ctrlGen.Caption
'            If (strDsc <> "") Then strDsc = ctrlGen.Text
'    End Select
'    On Local Error GoTo 0
'    CtrlGetDsc = strNom
'    If (strDsc <> "") Then CtrlGetDsc = CtrlGetDsc & " '" & strDsc & "'"
'End Function

Private Function DammiAgtGrp() As enmAgtGruppo
    Dim setAgtGrp As enmAgtGruppo

    setAgtGrp = agtNonEsegui
    If (opt(OPT_ESEGUI).Value) Then
        setAgtGrp = agtDopo
        If (optP(OPT_ESE_GRUPPO).Value) Then setAgtGrp = agtPrima
    End If
    DammiAgtGrp = setAgtGrp

End Function

Sub DefCampi()
    'definizione combo visualizzazione
'    cmbVis.addItem CStr(MXNU.CaricaStringaRes(75000)), visPerTipo
'    cmbVis.addItem CStr(MXNU.CaricaStringaRes(75001)), visPerScheda
End Sub

Sub DefLingua()
    Me.Caption = MXNU.CaricaStringaRes(23009)
'    pnlBack.Caption = MXNU.CaricaStringaRes(24028)
'    frmFrame(0).Caption = MXNU.CaricaStringaRes(24023)
'    frmFrame(1).Caption = MXNU.CaricaStringaRes(24024)
'    frmFrame(2).Caption = MXNU.CaricaStringaRes(24015)
'    ComClr.Caption = MXNU.CaricaStringaRes(25018)
'    com(COM_IMPOSTA).Caption = MXNU.CaricaStringaRes(25019)
'    com(COM_ANNULLA).Caption = MXNU.CaricaStringaRes(25008)
'    com(COM_COPIA).Caption = MXNU.CaricaStringaRes(25020)
    Call MXNU.LeggiRisorseControlli(Me)
End Sub

Sub Impostazioni_Copy()

    Dim vetParam() As Variant
    Dim strSrcKey As String, strSrcTip As String, strSrcDsc As String
    Dim strDstKey As String, strDstTip As String, strDstDsc As String
    Dim CAgtSrc As CAgeUtenti, CAgtDst As CAgeUtenti

    strSrcTip = Left$(trwUtenti.SelectedItem.tag, 1)
    strSrcKey = Mid$(trwUtenti.SelectedItem.tag, 2)
    strSrcDsc = trwUtenti.SelectedItem.text
    If (frmSelUtente.SelezionaGruppoUtente(opeCopiaAccessi, strDstKey, strDstTip, strDstDsc, True)) Then
        ReDim vetParam(1 To 4) As Variant
        If (strSrcTip = "G") Then vetParam(1) = MXNU.CaricaStringaRes(24029) Else vetParam(1) = MXNU.CaricaStringaRes(24030)
        vetParam(2) = strSrcDsc
        If (strDstTip = "G") Then vetParam(3) = MXNU.CaricaStringaRes(24029) Else vetParam(3) = MXNU.CaricaStringaRes(24030)
        vetParam(4) = strDstDsc
        If (MsgBox(MXNU.CaricaStringaRes(1036, vetParam()), vbQuestion + vbYesNo) = vbYes) Then
            If (strSrcTip = "G") Then Set CAgtSrc = colAgtGruppi(strSrcTip & strSrcKey) Else Set CAgtSrc = colAgtUtenti(strSrcTip & strSrcKey)
            If (strDstTip = "G") Then Set CAgtDst = colAgtGruppi(strDstTip & strDstKey) Else Set CAgtDst = colAgtUtenti(strDstTip & strDstKey)
            Call CAgtDst.CopiaImpostazioni(CAgtSrc)
            Set CAgtSrc = Nothing
            Set CAgtDst = Nothing
        End If
    End If

End Sub

Private Sub Impostazioni_Riporta(strAgente As String, setAgtGrp As enmAgtGruppo)
    'imposto i controlli
    setStato = stsReadOnly
    Call pnlBack_Enable(True)
    cmbAgenti.ListIndex = CercaAgente(strAgente)
    If (setAgtGrp = agtNonEsegui) Then
        opt(OPT_OMETTI).Value = True
    Else
        opt(OPT_ESEGUI).Value = True
        optP(OPT_ESE_GRUPPO).Value = (setAgtGrp = agtPrima)
        optP(OPT_ESE_UTENTE).Value = (setAgtGrp = agtDopo)
    End If
    setStato = stsWrite
End Sub

Sub Impostazioni_Write(ByVal nodUtente As MSComctlLib.Node, _
                        strAgente As String, _
                        setAgtGrp As enmAgtGruppo)
    Dim cNode As MSComctlLib.Node

    If Not (nodUtente Is Nothing) Then
        If (nodUtente <> trwUtenti.Nodes("Metodo98") And nodUtente <> trwUtenti.Nodes("gruppi")) Then
            If (Left$(nodUtente.tag, 1) = "G") Then
                'memorizzo impostazioni gruppo
                Call colAgtGruppi(CStr(nodUtente.tag)).MemorizzaImpostazioni(strAgente, setAgtGrp)
            Else
                'memorizzo impostazioni utente
                If nodUtente = trwUtenti.Nodes("utenti") Then
                    'valido per tutti gli utenti
                    If nodUtente.children > 0 Then
                        Set cNode = nodUtente.Child
                        Do While Not (cNode Is Nothing)
                            Call colAgtUtenti(CStr(cNode.tag)).MemorizzaImpostazioni(strAgente, setAgtGrp)
                            Set cNode = cNode.Next
                        Loop
                    End If
                    Set cNode = Nothing
                Else
                    Call colAgtUtenti(CStr(nodUtente.tag)).MemorizzaImpostazioni(strAgente, setAgtGrp)
                End If
            End If
        End If
    End If

End Sub

'Function KeyFormGet(frmDef As Form, setEvento As enmAgtEventi) As String
'    KeyFormGet = frmDef.Name & "_" & setEvento
'End Function

Private Sub pnlBack_Refresh(ByVal nodUtente As MSComctlLib.Node)
    If (Left$(nodUtente.tag, 1) = "G") Then
        'nodo gruppo -> nascondo impostazioni agente gruppo
        opt(OPT_OMETTI).Visible = False
        opt(OPT_OMETTI).Value = False
        opt(OPT_ESEGUI).Visible = False
        opt(OPT_ESEGUI).Value = False
        optP(OPT_ESE_GRUPPO).Visible = False
        optP(OPT_ESE_GRUPPO).Value = False
        optP(OPT_ESE_UTENTE).Visible = False
        optP(OPT_ESE_UTENTE).Value = False
    Else
        'nodo gruppo -> mostro impostazioni agente gruppo
        opt(OPT_OMETTI).Visible = True
        opt(OPT_ESEGUI).Visible = True
        optP(OPT_ESE_GRUPPO).Visible = True
        optP(OPT_ESE_UTENTE).Visible = True
    End If
End Sub

Sub SalvaImpostazioni()

    Dim CImp As CAgeUtenti

    Screen.MousePointer = vbHourglass
    'salvo le impostazioni per i gruppi
    For Each CImp In colAgtGruppi
        Call CImp.SalvaImpostazioniGruppo(MlngIDForm)
    Next
    'salvo le impostazioni per gli utenti
    For Each CImp In colAgtUtenti
        Call CImp.SalvaImpostazioniUtente(MlngIDForm)
    Next
    Screen.MousePointer = vbDefault

End Sub

'Sub TreeOggetti_Add(ctrGen As Control, vntKeyParent As String)
'
'    Dim nodX As Node
'
'    Select Case TypeName(ctrGen)
'        Case "MWEtichetta", "MWSchedaBox"
'        Case "MWLinguetta"
'            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_LING)
'            Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'        Case "TextBox"
'            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_TEXTBOX)
'            Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'        Case "MaskEdBox"
'            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_MASKEDIT)
'            Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'        Case "CommandButton"
'            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_BUTTON)
'            Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'        Case "CheckBox"
'            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_CHECK)
'            Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'        Case "OptionButton"
'            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_OPTION)
'            Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'        Case "ComboBox"
'            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_COMBO)
'            Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'        Case "fpSpread"
'            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_SPREAD)
'            Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'        Case Else
'            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_ALTRE)
'            Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'    End Select
'
'End Sub
'
'Sub TreeOggetti_GetChild(ctrParent As Object, ByVal strKeyParent As String)
'Dim nodX As Node
'Dim ctrGen As Control
'
'    For Each ctrGen In frmDef
'        If (TypeName(ctrGen) <> "MWLinguetta") Then
'            If (ctrGen.Container.hwnd = ctrParent.hwnd) Then
'                Call TreeOggetti_Add(ctrGen, strKeyParent)
'            End If
'        End If
'    Next
'
'End Sub

'Sub TreeOggetti_Inizializza(setTipoVis As enmCtrlTipoVisione)
'    Dim nodX As Node
'    Dim ctrGen As Control
'
'    'imposto nodo form
'    Set nodX = trwOggetti.Nodes.Add(, tvwFirst, frmDef.Name, frmDef.Caption, "form")
'    nodX.Tag = "" 'non collegabile
'    Call nodX.EnsureVisible
'    '... e i suoi eventi
'    Set nodX = trwOggetti.Nodes.Add(frmDef.Name, tvwChild, KeyFormGet(frmDef, evtSalvaInserimento), "Inserimento", "evento")
'    Call AssegnaTag(nodX, frmDef.Name, evtSalvaInserimento)
'    Set nodX = trwOggetti.Nodes.Add(frmDef.Name, tvwChild, KeyFormGet(frmDef, evtSalvaModifica), "Modifica", "evento")
'    Call AssegnaTag(nodX, frmDef.Name, evtSalvaModifica)
'    Set nodX = trwOggetti.Nodes.Add(frmDef.Name, tvwChild, KeyFormGet(frmDef, evtAnnullamento), "Annullamento", "evento")
'    Call AssegnaTag(nodX, frmDef.Name, evtAnnullamento)
'    Set nodX = trwOggetti.Nodes.Add(frmDef.Name, tvwChild, KeyFormGet(frmDef, evtNuovo), "Nuovo Codice", "evento")
'    Call AssegnaTag(nodX, frmDef.Name, evtNuovo)
'
'    If (setTipoVis = visPerTipo) Then
'        '>>>VISIONE PER TIPO CONTROLLO
'        'radice linguette
'        Set nodX = trwOggetti.Nodes.Add(, tvwNext, KEY_LING, "Linguette", PIC_CTRL)
'        Call nodX.EnsureVisible
'        nodX.Sorted = True
'        nodX.Tag = "" 'non collegabile
'        'radice textbox
'        Set nodX = trwOggetti.Nodes.Add(, tvwNext, KEY_TEXTBOX, "Campi di Testo", PIC_CTRL)
'        nodX.Sorted = True
'        Call nodX.EnsureVisible
'        nodX.Tag = "" 'non collegabile
'        'radice mask edit
'        Set nodX = trwOggetti.Nodes.Add(, tvwNext, KEY_MASKEDIT, "Campi Data", PIC_CTRL)
'        nodX.Sorted = True
'        Call nodX.EnsureVisible
'        nodX.Tag = "" 'non collegabile
'        'radice bottoni
'        Set nodX = trwOggetti.Nodes.Add(, tvwNext, KEY_BUTTON, "Bottoni", PIC_CTRL)
'        nodX.Sorted = True
'        Call nodX.EnsureVisible
'        nodX.Tag = "" 'non collegabile
'        'radice check box
'        Set nodX = trwOggetti.Nodes.Add(, tvwNext, KEY_CHECK, "Campi di Selezione", PIC_CTRL)
'        nodX.Sorted = True
'        Call nodX.EnsureVisible
'        nodX.Tag = "" 'non collegabile
'        'radice option
'        Set nodX = trwOggetti.Nodes.Add(, tvwNext, KEY_OPTION, "Campi Opzioni", PIC_CTRL)
'        nodX.Sorted = True
'        Call nodX.EnsureVisible
'        nodX.Tag = "" 'non collegabile
'        'radice combo box
'        Set nodX = trwOggetti.Nodes.Add(, tvwNext, KEY_COMBO, "Liste Selezione", PIC_CTRL)
'        nodX.Sorted = True
'        Call nodX.EnsureVisible
'        nodX.Tag = "" 'non collegabile
'        'radice fogli
'        Set nodX = trwOggetti.Nodes.Add(, tvwNext, KEY_SPREAD, "Fogli di Calcolo", PIC_CTRL)
'        nodX.Sorted = True
'        Call nodX.EnsureVisible
'        nodX.Tag = "" 'non collegabile
'        'radice altri
'        Set nodX = trwOggetti.Nodes.Add(, tvwNext, KEY_ALTRE, "Altri Controlli", PIC_CTRL)
'        nodX.Sorted = True
'        Call nodX.EnsureVisible
'        nodX.Tag = "" 'non collegabile
'
'        For Each ctrGen In frmDef
'            Select Case TypeName(ctrGen)
'                Case "MWLinguetta"
'                    Call TreeOggetti_Add(ctrGen, KEY_LING)
'                Case "TextBox"
'                    Call TreeOggetti_Add(ctrGen, KEY_TEXTBOX)
'                Case "MaskEdBox"
'                    Call TreeOggetti_Add(ctrGen, KEY_MASKEDIT)
'                Case "CommandButton"
'                    Call TreeOggetti_Add(ctrGen, KEY_BUTTON)
'                Case "CheckBox"
'                    Call TreeOggetti_Add(ctrGen, KEY_CHECK)
'                Case "OptionButton"
'                    Call TreeOggetti_Add(ctrGen, KEY_OPTION)
'                Case "ComboBox"
'                    Call TreeOggetti_Add(ctrGen, KEY_COMBO)
'                Case "fpSpread"
'                    Call TreeOggetti_Add(ctrGen, KEY_SPREAD)
'                Case Else
'                    Call TreeOggetti_Add(ctrGen, KEY_ALTRE)
'            End Select
'        Next ctrGen
'    Else
'        '>>>VISIONE PER SCHEDA
'        'per ogni linguetta...
'        For Each ctrGen In frmDef
'            If (TypeName(ctrGen) = "MWLinguetta" And StrComp(ctrGen.Name, "Ling", vbTextCompare) = 0) Then
'                '...imposto i controlli contenuti nella scheda ad essa collegata
'                Set nodX = trwOggetti.Nodes.Add(, tvwNext, KeyControlloGet(ctrGen), swapp(ctrGen.Caption, "&", ""), "ling")
'                Call nodX.EnsureVisible
'                Call AssegnaTag(nodX, KeyControlloGet(ctrGen), evtGenerico)
'                'cerca eventuali gli oggetti contenuti
'                Call TreeOggetti_GetChild(frmDef.Scheda(ctrGen.Index), KeyControlloGet(ctrGen))
'            End If
'        Next
'    End If
'
'fine_Inizializza:
'    bolOggImp = True
'    Exit Sub
'
'End Sub

Private Sub cmbAgenti_Click()
    If (setStato = stsWrite) Then
        Call Impostazioni_Write(trwUtenti.SelectedItem, cmbAgenti.text, DammiAgtGrp())
    End If
End Sub

'Private Sub cmbVis_Click()
'
'    trwOggetti.Nodes.Clear
'    Call TreeOggetti_Inizializza(cmbVis.ListIndex)
'
'End Sub


Private Sub com_Click(Index As Integer)
    Select Case Index
        Case COM_IMPOSTA
            Call SalvaImpostazioni
            Unload Me
        Case COM_ANNULLA
            Unload Me
        Case COM_COPIA
           Call Impostazioni_Copy
    End Select
End Sub


Private Sub Form_Initialize()
    'Set colOggetti = New Collection
    Set colAgtUtenti = New Collection
    Set colAgtGruppi = New Collection
End Sub

Private Sub Form_KeyPress(keyAscii As Integer)
    If (keyAscii = vbKeyEscape) Then Unload Me
End Sub

Private Sub Form_Load()
    If (frmDef.HelpContextID <> 0) Then
        bolOggImp = False
        setStato = stsWrite
        MlngIDForm = frmDef.HelpContextID
        etcfinestra.Caption = frmDef.Caption
        Call DefCampi
        Call DefLingua
        'carico combo agenti
        Call Agenti_CaricaCombo
        Call TreeUtenti_Inizializza(trwUtenti, True, True)
        'cmbVis.ListIndex = visPerTipo
        Call InizializzaStrutture
        Call trwUtenti_NodeClick(trwUtenti.Nodes("Metodo98"))
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
        'mostro la finestra
        Call CentraFinestra(Me.hwnd)
    Else
        MsgBox MXNU.CaricaStringaRes(1040), vbCritical
        Unload Me
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set colOggetti = Nothing
    Set colAgtUtenti = Nothing
    Set colAgtGruppi = Nothing

    Set frmDefAge = Nothing
End Sub


Sub InizializzaStrutture()

    Dim intq As Integer
    Dim strSQL As String
    Dim hSS As CRecordSet
    Dim bolEnd As Boolean
    Dim vntAus As Variant
    Dim CAgt As CAgeUtenti

    'inizializzo struttua gruppi
    strSQL = "SELECT Codice FROM TabGruppiUtente"

    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    Do While (Not bolEnd)
        vntAus = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Codice", "")
        If (vntAus <> "") Then
            Set CAgt = New CAgeUtenti
            Call CAgt.Inizializza(vntAus, tipGruppo, MlngIDForm)
            Call colAgtGruppi.Add(CAgt, "G" & vntAus)
        End If
        bolEnd = Not MXDB.dbSuccessivo(hSS, TIPO_SNAPSHOT)
    Loop
    intq = MXDB.dbChiudiSS(hSS)

    'inizializzo struttua utenti
    strSQL = "SELECT UserID" _
            & " FROM TabUtenti"
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    Do While (Not bolEnd)
        vntAus = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "UserID", "")
        If (vntAus <> "") Then
            Set CAgt = New CAgeUtenti
            Call CAgt.Inizializza(vntAus, tipUtente, MlngIDForm)
            Call colAgtUtenti.Add(CAgt, "U" & vntAus)
        End If
        bolEnd = Not MXDB.dbSuccessivo(hSS, TIPO_SNAPSHOT)
    Loop
    intq = MXDB.dbChiudiSS(hSS)
End Sub

Private Sub opt_Click(Index As Integer)

    Exit Sub
    If (opt(OPT_OMETTI).Value) Then
        optP(OPT_ESE_GRUPPO).Enabled = False
        optP(OPT_ESE_UTENTE).Enabled = False
        optP(OPT_ESE_GRUPPO).Value = False
        optP(OPT_ESE_UTENTE).Value = False
    Else
        optP(OPT_ESE_GRUPPO).Enabled = True
        optP(OPT_ESE_UTENTE).Enabled = True
    End If

    If (setStato = stsWrite) Then
        Call Impostazioni_Write(trwUtenti.SelectedItem, cmbAgenti.text, DammiAgtGrp())
    End If

End Sub

'Private Sub TreeOggetti_Refresh(ByVal nodUtente As MSComctlLib.Node)
'    Dim nodX As Node
'
'    If (nodUtente Is Nothing) Then Exit Sub
'    If (nodUtente.key = "Metodo98" Or nodUtente.key = "gruppi" Or nodUtente.key = "utenti") Then
'        'disabilito l'albero degli oggetti
'        trwOggetti.Enabled = False
'        Call pnlBack_Enable(False)
'    Else
'        'trwOggetti.Enabled = True
'        'If (trwOggetti.SelectedItem Is Nothing) Then
'            'seleziono il nodo della finestra
'        '    trwOggetti.SelectedItem = trwOggetti.Nodes(CStr(frmDef.Name))
'        'End If
'        'Call trwOggetti_NodeClick(trwOggetti.SelectedItem)
'    End If
'
'End Sub

Private Sub pnlBack_Enable(bolEnabled As Boolean)
Dim cnt As Integer
    Exit Sub
    'disabilito il pannello...
    pnlBack.Enabled = bolEnabled
    '... e tutti i controlli in esso contenuti
    For cnt = OPT_OMETTI To OPT_ESEGUI
        opt(cnt).Enabled = bolEnabled
    Next cnt
    For cnt = OPT_ESE_GRUPPO To OPT_ESE_UTENTE
        optP(cnt).Enabled = bolEnabled
    Next cnt
    DoEvents
End Sub


Function Impostazioni_Read(ByVal nodUtente As MSComctlLib.Node, _
                            strAgente As String, _
                            setAgtGruppo As enmAgtGruppo) As Boolean

    If (nodUtente Is Nothing) Then
        Impostazioni_Read = False
    Else
        If (nodUtente.key = "Metodo98" Or nodUtente.key = "gruppi" Or nodUtente.key = "utenti") Then
            Impostazioni_Read = False
        Else
            Impostazioni_Read = True
            If (Left$(nodUtente.tag, 1) = "G") Then
                'Agenti gruppi
                Call colAgtGruppi(CStr(nodUtente.tag)).LeggiImpostazioni(strAgente, setAgtGruppo)
            Else
                'Agenti utente
                Call colAgtUtenti(CStr(nodUtente.tag)).LeggiImpostazioni(strAgente, setAgtGruppo)
            End If
        End If
    End If

End Function


Private Sub OptP_Click(Index As Integer)
    If (setStato = stsWrite) Then
        Call Impostazioni_Write(trwUtenti.SelectedItem, cmbAgenti.text, DammiAgtGrp())
    End If
End Sub




Private Sub Scheda_Paint()
    Call SchedaOmbreggiaControlli(Scheda())
End Sub

'Private Sub trwOggetti_NodeClick(ByVal Node As MSComctlLib.Node)
'Dim strAgente As String, setAgtGrp As enmAgtGruppo
'    setStato = stsReadOnly
'    If (Node.Tag = "") Then
'        Call pnlBack_Enable(False)
'    Else
'        'leggo le impostazioni per il nodo selezionato
'        If (Impostazioni_Read(trwUtenti.SelectedItem, Node, strAgente, setAgtGrp)) Then
'            Call Impostazioni_Riporta(strAgente, setAgtGrp)
'        End If
'    End If
'    setStato = stsWrite
'End Sub

Private Sub trwUtenti_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim strAgente As String, setAgtGrp As enmAgtGruppo

    setStato = stsReadOnly

    'refresh dell'albero degli oggetti
'    Call TreeOggetti_Refresh(Node)
    Call pnlBack_Refresh(Node)
    If (Impostazioni_Read(Node, strAgente, setAgtGrp)) Then
        Call Impostazioni_Riporta(strAgente, setAgtGrp)
    End If
    'abilito/disabilito il bottone copia
    com(COM_COPIA).Enabled = (Node.key <> "Metodo98" And Node.key <> "gruppi" And Node.key <> "utenti")
    setStato = stsWrite
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

