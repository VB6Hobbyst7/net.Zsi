VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F052EB6D-0FEB-4155-A1F2-38B5F18F06AD}#1.1#0"; "mxctrl.ocx"
Begin VB.Form FrmSchedula 
   ClientHeight    =   5295
   ClientLeft      =   1965
   ClientTop       =   1725
   ClientWidth     =   9435
   Icon            =   "Schedula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9435
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   5325
      Index           =   2
      Left            =   -90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   9393
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleWidth      =   9555
      ScaleHeight     =   5325
      Begin VB.ComboBox Cbo 
         Height          =   315
         Index           =   3
         Left            =   7425
         TabIndex        =   42
         Text            =   "Cbo3"
         Top             =   2925
         Visible         =   0   'False
         Width           =   1770
      Appearance      =   0  'Flat
      BeginProperty Font
      Weight          =   400
      EndProperty  
      End
      Begin VB.TextBox TxtBox 
      Height          =   270
         Index           =   2
         Left            =   7965
         TabIndex        =   41
         Text            =   "txtbox(2)"
         top = 1995
         Visible         =   0   'False
         Width           =   1245
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font
      Weight          =   400
      EndProperty  
      End
      Begin VB.TextBox TxtBox 
      Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   7830
         PasswordChar    =   "*"
         TabIndex        =   40
         Text            =   "txtbox(1)"
         top = 1635
         Visible         =   0   'False
         Width           =   1260
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font
      Weight          =   400
      EndProperty  
      End
      Begin VB.TextBox TxtBox 
      Height          =   270
         Index           =   0
         Left            =   7830
         TabIndex        =   39
         Text            =   "txtbox(0)"
         top = 1140
         Visible         =   0   'False
         Width           =   1245
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font
      Weight          =   400
      EndProperty  
      End
      Begin VB.TextBox TxtBox 
      Height          =   270
         Index           =   5
         Left            =   7605
         TabIndex        =   38
         Text            =   "txtbox(5)"
         top = 285
         Visible         =   0   'False
         Width           =   1305
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font
      Weight          =   400
      EndProperty  
      End
      Begin VB.ComboBox Cbo 
         Height          =   315
         Index           =   0
         Left            =   7425
         TabIndex        =   37
         Text            =   "Cbo"
         Top             =   720
         Visible         =   0   'False
         Width           =   1725
      Appearance      =   0  'Flat
      BeginProperty Font
      Weight          =   400
      EndProperty  
      End
      Begin VB.ComboBox Cbo 
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   30
         Text            =   "Cbo"
         Top             =   540
         Width           =   3975
      Appearance      =   0  'Flat
      BeginProperty Font
      Weight          =   400
      EndProperty  
      End
      Begin VB.Frame Fra 
         Height          =   960
         Index           =   1
         Left            =   180
         TabIndex        =   27
         Top             =   1710
         Width           =   2265
         Appearance      =   0  'Flat
         BorderStyle     =   1  'None
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Begin VB.TextBox TxtBox 
         Height          =   270
            Index           =   7
            Left            =   225
            MaxLength       =   4
            TabIndex        =   28
            Text            =   "txtbox(7)"
            top = 465
            Width           =   1530
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font
         Weight          =   400
         EndProperty  
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Ripetizione ogni tot minuti"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   29
            Top             =   180
            WhatsThisHelpID =   11068
            Width           =   1800
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         End
      End
      Begin VB.Frame Fra 
         Height          =   1815
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   2655
         Width           =   2670
         Appearance      =   0  'Flat
         BorderStyle     =   1  'None
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Begin VB.CheckBox Chk 
            Caption         =   "Lun"
            Height          =   330
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   270
            WhatsThisHelpID =   50211
            Width           =   1050
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Mar"
            Height          =   330
            Index           =   1
            Left            =   1260
            TabIndex        =   25
            Top             =   270
            WhatsThisHelpID =   50212
            Width           =   1365
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Mer"
            Height          =   420
            Index           =   2
            Left            =   180
            TabIndex        =   24
            Top             =   585
            WhatsThisHelpID =   50213
            Width           =   1050
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Gio"
            Height          =   420
            Index           =   3
            Left            =   1260
            TabIndex        =   23
            Top             =   585
            WhatsThisHelpID =   50214
            Width           =   1300
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Ven"
            Height          =   420
            Index           =   4
            Left            =   180
            TabIndex        =   22
            Top             =   945
            WhatsThisHelpID =   50215
            Width           =   1005
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Sab"
            Height          =   420
            Index           =   5
            Left            =   1260
            TabIndex        =   21
            Top             =   945
            WhatsThisHelpID =   50216
            Width           =   1320
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Dom"
            Height          =   420
            Index           =   6
            Left            =   1260
            TabIndex        =   20
            Top             =   1350
            WhatsThisHelpID =   50217
            Width           =   1320
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
      End
      Begin VB.Frame Fra 
         Height          =   1815
         Index           =   4
         Left            =   2880
         TabIndex        =   6
         Top             =   2655
         Width           =   4065
         Appearance      =   0  'Flat
         BorderStyle     =   1  'None
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Begin VB.CheckBox Chk 
            Caption         =   "gen"
            Height          =   330
            Index           =   7
            Left            =   180
            TabIndex        =   18
            Top             =   315
            WhatsThisHelpID =   50351
            Width           =   915
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "feb"
            Height          =   330
            Index           =   8
            Left            =   180
            TabIndex        =   17
            Top             =   630
            WhatsThisHelpID =   50352
            Width           =   915
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "mar"
            Height          =   420
            Index           =   9
            Left            =   180
            TabIndex        =   16
            Top             =   900
            WhatsThisHelpID =   50353
            Width           =   780
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "apr"
            Height          =   420
            Index           =   10
            Left            =   180
            TabIndex        =   15
            Top             =   1215
            WhatsThisHelpID =   50354
            Width           =   870
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "mag"
            Height          =   420
            Index           =   11
            Left            =   1260
            TabIndex        =   14
            Top             =   270
            WhatsThisHelpID =   50355
            Width           =   1095
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "giu"
            Height          =   420
            Index           =   12
            Left            =   1260
            TabIndex        =   13
            Top             =   585
            WhatsThisHelpID =   50356
            Width           =   1095
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "lug"
            Height          =   375
            Index           =   13
            Left            =   1260
            TabIndex        =   12
            Top             =   900
            WhatsThisHelpID =   50357
            Width           =   1005
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "ago"
            Height          =   330
            Index           =   14
            Left            =   1260
            TabIndex        =   11
            Top             =   1215
            WhatsThisHelpID =   50358
            Width           =   1140
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "set"
            Height          =   420
            Index           =   15
            Left            =   2475
            TabIndex        =   10
            Top             =   270
            WhatsThisHelpID =   50359
            Width           =   1500
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "otto"
            Height          =   420
            Index           =   16
            Left            =   2475
            TabIndex        =   9
            Top             =   585
            WhatsThisHelpID =   50360
            Width           =   1455
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "nove"
            Height          =   420
            Index           =   17
            Left            =   2475
            TabIndex        =   8
            Top             =   900
            WhatsThisHelpID =   50361
            Width           =   1455
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
         Begin VB.CheckBox Chk 
            Caption         =   "dic"
            Height          =   420
            Index           =   18
            Left            =   2475
            TabIndex        =   7
            Top             =   1170
            WhatsThisHelpID =   50362
            Width           =   1455
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         End
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "Schedla...."
         Height          =   375
         Index           =   0
         Left            =   7155
         TabIndex        =   5
         Top             =   4770
         WhatsThisHelpID =   25095
         Width           =   2040
      End
      Begin VB.TextBox TxtBox 
      Height          =   270
         Index           =   6
         Left            =   5535
         TabIndex        =   4
         Text            =   "txtbox(6)"
         top = 870
         Visible         =   0   'False
         Width           =   1260
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font
      Weight          =   400
      EndProperty  
      End
      Begin VB.Frame Fra 
         Height          =   960
         Index           =   5
         Left            =   2880
         TabIndex        =   1
         Top             =   1710
         Width           =   1905
         Appearance      =   0  'Flat
         BorderStyle     =   1  'None
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Begin VB.ComboBox Cbo 
            Height          =   315
            Index           =   2
            Left            =   135
            TabIndex        =   2
            Text            =   "Cbo(2)"
            Top             =   495
            Width           =   1500
         Appearance      =   0  'Flat
         BeginProperty Font
         Weight          =   400
         EndProperty  
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Giorno ddel mese!!!"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   3
            Top             =   225
            WhatsThisHelpID =   11069
            Width           =   1365
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         End
      End
      Begin MSMask.MaskEdBox MskEd 
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   31
         top = 1230
         WhatsThisHelpID =   1096
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      BorderStyle     =   0
      End
      Begin MSMask.MaskEdBox MskEd 
         Height          =   270
         Index           =   1
         Left            =   1665
         TabIndex        =   32
         top = 1230
         WhatsThisHelpID =   1096
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   " "
      BorderStyle     =   0
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Ora Iniziooo"
         Height          =   195
         Index           =   3
         Left            =   1665
         TabIndex        =   36
         Top             =   990
         WhatsThisHelpID =   11067
         Width           =   840
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Data Iniziooo"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   35
         Top             =   945
         WhatsThisHelpID =   11066
         Width           =   930
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Qu"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   34
         Top             =   270
         WhatsThisHelpID =   11065
         Width           =   210
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "id operazione......"
         Height          =   195
         Index           =   13
         Left            =   5625
         TabIndex        =   33
         Top             =   585
         Visible         =   0   'False
         WhatsThisHelpID =   11065
         Width           =   1215
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      End
   End
End
Attribute VB_Name = "FrmSchedula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Dim WithEvents xAna As MXKit.Anagrafica
Attribute xAna.VB_VarHelpID = -1
Private strMsg As String
Private mStrTipoOperazione As String
Private mColParametriOP As Collection

Private mIntTipoOperazione As Integer

Dim mBolSalvata As Boolean


Public Sub CaricaarrayGM(Giorni As Variant, Mesi As Variant)
ReDim Giorni(1 To 7)
ReDim Mesi(1 To 12)
Giorni(1) = "LUN"
Giorni(2) = "MAR"
Giorni(3) = "MER"
Giorni(4) = "GIO"
Giorni(5) = "VEN"
Giorni(6) = "SAB"
Giorni(7) = "DOM"

Mesi(1) = "GENN"
Mesi(2) = "FEBR"
Mesi(3) = "MARZ"
Mesi(4) = "APRI"
Mesi(5) = "MAGG"
Mesi(6) = "GIUG"
Mesi(7) = "LUGL"
Mesi(8) = "AGOS"
Mesi(9) = "SETT"
Mesi(10) = "OTTO"
Mesi(11) = "NOVE"
Mesi(12) = "DICE"
End Sub
Private Function ctrlSalvaRecord() As Boolean


  If ControllaImpostazioni Then
        ctrlSalvaRecord = True
  Else
        Call MXNU.MsgBoxEX(strMsg, vbCritical, 1007)
        ctrlSalvaRecord = False
  End If
End Function

Private Function SalvaRecord(DBSCH As adodb.Connection) As Boolean
Dim setResSalva As enmSalvaAnagrafica
Dim IDope As Variant
Dim DataProssimaEsecuzione As Date
Dim StrApp As String
Dim StrApp1 As String


    SalvaRecord = False
    If (ctrlSalvaRecord()) Then
       ' IDope = xAna.GrInput("IDOP").ValoreCorrente
      '  StrApp = xAna.GrInput("datainizio").ValoreCorrente
      '  StrApp1 = xAna.GrInput("orainizio").ValoreCorrente
      '  StrApp = Left(StrApp, 10) & " " & Right(StrApp1, 8)
      '  DataProssimaEsecuzione = StrApp
        
        
        Dim DBMET As adodb.Connection
        Set DBMET = hndDBArchivi.ConnessioneR
        Set hndDBArchivi.ConnessioneR = DBSCH
        
        setResSalva = xAna.SalvaAnagrafica()
        
        Set hndDBArchivi.ConnessioneR = DBMET
        Set DBMET = Nothing
        
        Select Case setResSalva
            Case saRegCorretta
                    SalvaRecord = True
            
            Case saRegFallita
                SalvaRecord = False
            
            Case saRegNuovoCodice
                If xAna.AssegnaNuovoCodice() Then
                    On Local Error Resume Next
                    Cbo(1).SetFocus
                    On Local Error GoTo 0
                End If
        End Select
    End If
End Function

Public Function Inizializza() As Boolean
    
    Inizializza = False
    
    On Local Error GoTo err_Inizializza
    'MstrFilelog = MXNU.PercorsoPreferenze & "\" & NOMEFILELOG
    
    Inizializza = True
fine_inizializza:
    On Local Error GoTo 0

Exit Function

err_Inizializza:
    Call MXNU.MsgBoxEX(Err.Description, vbCritical, "Inizializzazione Schedulatore - Apertura Database")
    Resume fine_inizializza
End Function

Private Sub Cbo_Click(Index As Integer)
    Dim strRifCtrl As String
    Dim TipoOp As String
    

    Select Case Index
        Case 0
            With xAna
                strRifCtrl = .NomeControllo2NomeVariabile("cbo_" & Index)
                Call .AssegnaCampo(strRifCtrl, Cbo(Index).ListIndex)
            End With
        Case 1
            '0=giorni
            '1=mesi
            '2=una sola volta
            '3=ogni tot minuti
            With xAna
                strRifCtrl = .NomeControllo2NomeVariabile("cbo_" & Index)
                Call .AssegnaCampo(strRifCtrl, Cbo(Index).ListIndex)
            End With

            Call ImpostaControlliTempi(Cbo(Index).ListIndex)
        Case 3
            With xAna
                Call .AssegnaCampo("ditta", Cbo(Index).text)
            End With
        Case Else
            With xAna
                strRifCtrl = .NomeControllo2NomeVariabile("cbo_" & Index)
                Call .AssegnaCampo(strRifCtrl, Cbo(Index).ListIndex)
            End With
       End Select
End Sub
Public Sub ImpostaControlliTempi(Valore As Integer)
        '0=giorni
        '1=mesi
        '2=una sola volta
        '3=ogni tot minuti
        Call AZZERACONTROLLI
        Fra(0).Visible = True
        Fra(1).Visible = True
        Fra(4).Visible = True
        Fra(5).Visible = True
        
       Select Case Valore
              
              Case 0
                    Fra(1).Visible = False
                    Fra(4).Visible = False
                    Fra(5).Visible = False
              
              Case 1
                    Fra(0).Visible = False
                    Fra(1).Visible = False

              Case 2
                    Fra(0).Visible = False
                    Fra(1).Visible = False
                    Fra(4).Visible = False
                    Fra(5).Visible = False
              
              Case 3
                    Fra(0).Visible = False
                    Fra(4).Visible = False
                    Fra(5).Visible = False
       End Select

End Sub
Public Sub AZZERACONTROLLI()
    Dim ctl As Variant
    
    On Local Error Resume Next
    For Each ctl In Me.Controls
        If ctl.Container Is Fra(0) Or ctl.Container Is Fra(1) Or ctl.Container Is Fra(4) Or ctl.Container Is Fra(5) Then
            Select Case ctl.Name
                    Case "Cbo", "TxtBox"
                           ctl.text = 0
                    
                    Case "Chk"
                        ctl.Value = 0
            End Select
        End If
    Next
    
End Sub
Private Sub Cbo_KeyPress(Index As Integer, keyAscii As Integer)
Dim strRifCtrl As String
Select Case Index
        
        
        Case 0
                Call CtrlKey(keyAscii, xAna.GrInput("quando").TipoInput)
        
        Case 1
              keyAscii = 0
        Case 3
                Call CtrlKey(keyAscii, xAna.GrInput("ditta").TipoInput)
                
        Case Else
            With xAna
                strRifCtrl = .NomeControllo2NomeVariabile("cbo_" & Index)
                Call CtrlKey(keyAscii, .GrInput(strRifCtrl).TipoInput)
            End With
End Select

End Sub

Private Sub Chk_Click(Index As Integer)
Dim strRifCtrl As String
With xAna
    strRifCtrl = .NomeControllo2NomeVariabile("chk_" & Index)
    Call xAna.AssegnaCampo(strRifCtrl, Chk(Index).Value)
End With

End Sub

Private Sub chk_KeyPress(Index As Integer, keyAscii As Integer)
Dim strRifCtrl As String

    With xAna
        strRifCtrl = .NomeControllo2NomeVariabile("chk_" & Index)
        Call CtrlKey(keyAscii, xAna.GrInput(strRifCtrl).TipoInput)
    End With

End Sub

Private Sub Cmd_Click(Index As Integer)
Dim IDope As Variant
Dim Ret As Long
Dim NomePar As String
Dim Valore1 As String
Dim Valore2 As String
Dim Valore3 As String
Dim Valore4 As String
Dim Valore5 As String
Dim Tipo1 As String
Dim Tipo2 As String
Dim Tipo3 As String
Dim Tipo4 As String
Dim Tipo5 As String
Dim Posizione As Integer
Dim StrSqlParametri As String
Dim DBSCH As adodb.Connection
Dim StrDBUser As String
Dim StrDBPW As String
Dim Strconn As String
Dim cPar As CParametri

mBolSalvata = True
On Local Error GoTo ERR_SalvaPARAM
Select Case Index
        Case 0
                Set DBSCH = New adodb.Connection
                DBSCH.ConnectionTimeout = 30
                DBSCH.CommandTimeout = 180
                DBSCH.CursorLocation = 3
                
                DBSCH = MXNU.GetstrConnection("SCHEDULATOREMET")
                StrDBUser = "trm" & MXNU.NTerminale
                StrDBPW = MXNU.PasswordDB
                DBSCH.Open Strconn, StrDBUser, StrDBPW
                
                IDope = xAna.GrInput("IDOP").ValoreCorrente
                
                'SALVO TUTTI I PARAMETRI
                

                'SALVO L'ANAGRAFICA DELLE FREQUENZE e se va a buon fine, faccio il commit dei parametri
                If SalvaRecord(DBSCH) Then
                    'inizio la trsansazione
                    DBSCH.IsolationLevel = adXactSerializable
                    DBSCH.BeginTrans
                    
                    For Posizione = 1 To mColParametriOP.Count
                       Set cPar = mColParametriOP(Posizione)
                       StrSqlParametri = "INSERT INTO PARAMETRI (IDOP, IDPAR, POSIZIONE, FLAGRIGA, NOMEPAR, VALORE1 ,TIPO1," & _
                       " VALORE2 , TIPO2, VALORE3 , TIPO3, VALORE4 ,TIPO4,VALORE5 ,TIPO5, UTENTEMODIFICA, DATAMODIFICA)" & _
                       "VALUES (" & xAna.GrInput("IDOP").ValoreCorrente & ", " & Posizione & "," & Posizione & ", 2 , " & _
                       hndDBArchivi.FormatoSQL(cPar.NomePar, DB_TEXT) & ", " & hndDBArchivi.FormatoSQL(cPar.Valore1, DB_TEXT) & ", " & hndDBArchivi.FormatoSQL(cPar.Tipo1, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(cPar.Valore2, DB_TEXT) & ", " & hndDBArchivi.FormatoSQL(cPar.Tipo2, DB_TEXT) & _
                       "," & hndDBArchivi.FormatoSQL(cPar.Valore3, DB_TEXT) & ", " & hndDBArchivi.FormatoSQL(cPar.Tipo3, DB_TEXT) & _
                       "," & hndDBArchivi.FormatoSQL(cPar.Valore4, DB_TEXT) & ", " & hndDBArchivi.FormatoSQL(cPar.Tipo4, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(cPar.Valore5, DB_TEXT) & ", " & hndDBArchivi.FormatoSQL(cPar.Tipo5, DB_TEXT) & ", '" & MXNU.UtenteAttivo & "', {Fn NOW()})"
                       DBSCH.Execute (StrSqlParametri)
                    Next Posizione
                    DBSCH.CommitTrans
                Else
                    mBolSalvata = False
                End If
                    
End Select

fine_salvaparam:
    DBSCH.Close
    Set DBSCH = Nothing
    'se salvataggio a buon fine -> nascondo la form
    If mBolSalvata Then
        Me.Hide
    End If
Exit Sub

ERR_SalvaPARAM:
        mBolSalvata = False
        If Err.Number <> 0 Then
             MXNU.MsgBoxEX Err.Number & " " & Err.Description, vbCritical, 1007
             'vado in rollback
             DBSCH.RollbackTrans
        End If
        If Err.Number <> 0 Then
            Resume fine_salvaparam
        End If
End Sub

Private Sub Command1_Click()
Dim DBditta As adodb.Connection
Dim Strconn As String
Dim StrDBUser As String
Dim StrDBPW As String
On Error GoTo err_prova
    Set DBditta = New adodb.Connection
    DBditta.ConnectionTimeout = 30
    DBditta.CommandTimeout = 180
    DBditta.CursorLocation = 3
    Strconn = MXNU.LeggiProfilo(MXNU.PercorsoLocal & "\DITTE.INI", "CONNESSIONE", xAna.GrInput("DITTA").ValoreCorrente, "")
    StrDBUser = xAna.GrInput("UTENTE").ValoreCorrente
    StrDBPW = xAna.GrInput("PASS").ValoreCorrente
    DBditta.Open Strconn, StrDBUser, StrDBPW
    DBditta.Close
    MXNU.MsgBoxEX 1470, vbOKOnly, 1007

err_prova_fine:

        Set DBditta = Nothing
Exit Sub


err_prova:
        MXNU.MsgBoxEX Err.Number & " " & Err.Description, vbCritical, 1007
        Resume err_prova_fine
End Sub

Private Sub Form_Load()
            Call InitControlli
            Call InitAna
            Call MXNU.LeggiRisorseControlli(Me)
            FrmSchedula.Caption = MXNU.CaricaStringaRes(23327)
End Sub
Public Sub InitControlli()
        Dim i As Integer
        Cbo(1).addItem MXNU.CaricaStringaRes(75351) '"giorni"
        Cbo(1).addItem MXNU.CaricaStringaRes(75352) ' "mesi"
        Cbo(1).addItem MXNU.CaricaStringaRes(75353) '"Una sola volta"
        Cbo(1).addItem MXNU.CaricaStringaRes(75354) '"ogni tot minuti"
        Cbo(1).text = Cbo(1).List(0)
        
        Cbo(2).addItem ""
        For i = 1 To 31
                Cbo(2).addItem i
        Next i
        Cbo(2).text = Cbo(2).List(0)
        
'        i = 1
'        While MXNU.LeggiProfilo(MXNU.PercorsoPgm & "\MWSchedula.ini", "TIPI OPERAZIONE", i, "") <> ""
'            Cbo(0).addItem MXNU.LeggiProfilo(MXNU.PercorsoPgm & "\MWSchedula.ini", "TIPI OPERAZIONE", i, "")
'            i = i + 1
'        Wend
'        Cbo(0).Text = ""
'        Cbo(0).ListIndex = -1
                
 '       i = 1
        Dim varLista As Variant
        Dim q As Integer
        Dim strRis As String
 '       strRis = MXNU.LeggiProfilo(MXNU.PercorsoPgm & "\Local\DITTE.ini", "DITTE", 0&, "")
 '       varLista = Split(strRis, Chr$(0), , vbTextCompare)
 '       For i = 0 To UBound(varLista) - 1
 '           Cbo(3).addItem varLista(i)
 '       Next i
'        Cbo(3).Text = ""
 '       Cbo(3).ListIndex = -1
        
        
End Sub
Public Sub InitAna()
Dim strEntry As String

    'inizzializzazione ana frequenza
    Set xAna = MXVA.CreaCAnagrafica("AnagraficaQuando", Me)
    Call xAna.Disegna
    xAna.GrInput("datainizio").Default = Date
    xAna.GrInput("orainizio").Formattazione = MXNU.Formato_HHMM
    xAna.GrInput("orainizio").Default = Now 'MXNU.Default_HHMM
    MskEd(1).Mask = MXNU.Mask_HHMM
    'subito entro in inserimento dell'anagrafica e setto il campo ricalcola=true
    xAna.Inserisci
    'ASSEGNO I VALORI ATTUALI DI UTENTE PASSWOR E ESERCIZIO
    Call xAna.AssegnaCampo("UTENTE", MXNU.UtenteAttivo)
    Call xAna.AssegnaCampo("PASS", MXNU.PasswordUtente)
    Call xAna.AssegnaCampo("ESERCIZIO", MXNU.AnnoAttivo)
    Call xAna.AssegnaCampo("DescrOP", TipoOperazione)
    Call xAna.AssegnaCampo("DITTA", MXNU.DittaAttiva)
    If RevLeggiProfilo(MXNU.PercorsoPgm & "\MWSchedula.ini", "TIPI OPERAZIONE", TipoOperazione, strEntry) Then
        Call xAna.AssegnaCampo("NOMEOP", Val(strEntry) - 1)
    End If
    'In questo modo dico allo schedulatore vero e proprio che è da ricalcolare la data di prosima esecuzione dell'operazione
    Call xAna.AssegnaCampo("Ricalcola", 1)
End Sub

Private Function RevLeggiProfilo(ByVal strFileIni As String, ByVal strSezione As String, ByVal strKey As String, strEntry As String) As Boolean
Dim vntDummy As Variant
Dim i As Integer
Dim bolFound As Boolean
Dim strMatch As String

    vntDummy = MXNU.LeggiProfilo(strFileIni, strSezione, 0&, "")
    If Len(vntDummy) > 0 Then
        vntDummy = Split(vntDummy, vbNullChar)
        bolFound = False
        For i = 0 To UBound(vntDummy)
            strMatch = MXNU.LeggiProfilo(strFileIni, strSezione, vntDummy(i), "")
            If StrComp(strMatch, strKey, vbTextCompare) = 0 Then
                strEntry = vntDummy(i)
                bolFound = True
                Exit For
            End If
        Next i
    End If
    RevLeggiProfilo = bolFound
End Function


Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    Set xAna = Nothing
End Sub


Public Function ControllaImpostazioni()
    Dim MESE(12) As Integer
    Dim GGmese As Integer
    Dim Giorno As Integer
    Dim num As Integer
    Dim i As Integer
    Dim k As Integer
    Dim ora As String
    Dim Minu As String
    Dim conv As Variant
    Dim BoolAlmenoUno As Boolean

    On Local Error GoTo Err_impostazioni
        '0=giorni
        '1=mesi
        '2=una sola volta
        '3=ogni tot minuti
ControllaImpostazioni = False
BoolAlmenoUno = False
strMsg = ""

If Len(Trim(MskEd(1).text)) >= 3 Then
    ora = Mid(MskEd(1).text, 1, InStr(1, MskEd(1).text, MXNU.Sep_time) - 1)
    ora = Right("00" & Trim(ora), 2)
    Minu = Mid(MskEd(1).text, 1 + InStr(1, MskEd(1).text, MXNU.Sep_time), Len(MskEd(1).text) - InStr(1, MskEd(1).text, MXNU.Sep_time))
    Minu = Right("00" & Trim(Minu), 2)
    conv = "2000-01-01 " & ora & MXNU.Sep_time & Minu & MXNU.Sep_time & "00"
    conv = CDate(conv)
Else
        strMsg = MXNU.CaricaStringaRes(1457)
        GoTo Err_impostazioni
End If

If Not IsDate(MskEd(0).text) Or Not IsDate(conv) Then
        strMsg = MXNU.CaricaStringaRes(1458)
        GoTo Err_impostazioni
End If

'>>CONTROLLO CHE I DATI di quando eseguire l'operazione siano corretti
Select Case Cbo(1).ListIndex
       Case 0 '0=giorni
                For i = 0 To 6
                    If Chk(i).Value = 1 Then
                        BoolAlmenoUno = True
                        Exit For
                    End If
                Next i
                If Not BoolAlmenoUno Then
                    strMsg = MXNU.CaricaStringaRes(1459)
                    GoTo Err_impostazioni
                End If
       Case 1 '1=mesi
                Giorno = Cbo(2).text
                If Not IsNumeric(Giorno) Then
                        strMsg = MXNU.CaricaStringaRes(1460)
                        GoTo Err_impostazioni
                End If
                k = 0
                For i = 1 To 12
                   If Chk(i + 6).Value = 1 Then
                       MESE(k) = i
                       k = k + 1
                   End If
                Next i
                
                If k = 0 Then
                     strMsg = MXNU.CaricaStringaRes(1461)
                     GoTo Err_impostazioni
                End If
                
                For i = 0 To k - 1
                    GGmese = MXNU.FineMese(Giorno & "/" & MESE(i) & "/" & Year(MskEd(0).text))
                    If (Giorno < 1) Or (Giorno > GGmese) Then
                        'StrMsg = "Impossibile eseguire l'operazione il giorno " & Giorno & " del mese " & MESE(i) & " (data non valida)"
                        strMsg = MXNU.CaricaStringaRes(1462, Array(Giorno, MESE(i)))
                        GoTo Err_impostazioni
                    End If
                Next i
       Case 2 '2=una sola volta
                    
       Case 3 '3=ogni tot minuti
                If Not IsNumeric(TxtBox(7).text) Or TxtBox(7).text <= 0 Then
                       strMsg = MXNU.CaricaStringaRes(1463)
                End If
                
       Case Else
             strMsg = MXNU.CaricaStringaRes(1464)
End Select

  

 'SE SONO ARRIVATO QUI SIGNIFICA CHE TUTTO è ANDATO BENE QUINDI RITORNO TRUE
 ControllaImpostazioni = True
Exit Function
Err_impostazioni:
                    If strMsg = "" Then
                        strMsg = Err.Number & " " & Err.Description
                    End If
                    


End Function
Private Sub mskEd_GotFocus(Index As Integer)
 SelContenuto MskEd(Index)
End Sub

Private Sub mskEd_KeyPress(Index As Integer, keyAscii As Integer)
Dim strRifCtrl As String
 ' If Index <> 1 Then
    With xAna
        strRifCtrl = .NomeControllo2NomeVariabile("msked_" & Index)
        Call CtrlKey(keyAscii, .GrInput(strRifCtrl).TipoInput)
    End With
 ' End If
End Sub

Private Sub mskEd_LostFocus(Index As Integer)
Dim strRifCtrl As String
Dim conv As String
Dim ora As String
Dim Minu As String

 
If Index = 1 Then
 On Local Error GoTo err_Imp
    If Len(MskEd(Index).text) > 2 Then
         ora = Mid(MskEd(Index).text, 1, InStr(1, MskEd(Index).text, MXNU.Sep_time) - 1)
         ora = Right("00" & Trim(ora), 2)
         Minu = Mid(MskEd(Index).text, 1 + InStr(1, MskEd(Index).text, MXNU.Sep_time), Len(MskEd(Index).text) - InStr(1, MskEd(Index).text, MXNU.Sep_time))
         Minu = Right("00" & Trim(Minu), 2)
         conv = "2000-01-01 " & ora & MXNU.Sep_time & Minu & MXNU.Sep_time & "00"
         conv = CDate(conv)
         With xAna
            If IsDate(conv) Then
                    .GrInput("orainizio").ValoreCorrente = conv
                    .GrInput("orainizio").Modificato = True
                    .flgModRecord = True
                    MskEd(Index).text = ora & MXNU.Sep_time & Minu
            Else
                    Call MXNU.MsgBoxEX(1466, vbExclamation, 1007)
                    MskEd(Index).SetFocus
            End If
            
         End With

    
    Else
           
           Call MXNU.MsgBoxEX(1466, vbExclamation, 1007)
           MskEd(Index).SetFocus
    End If
    On Local Error GoTo 0
 
 
 Else
        On Local Error Resume Next
        With xAna
          strRifCtrl = .NomeControllo2NomeVariabile("msked_" & Index)
          If MXNU.CtrlData(MskEd(Index)) Then
             If Err.Number = 0 Then
                  Call .AssegnaCampo(strRifCtrl, MskEd(Index).text)
              Else
                  MskEd(Index).text = .GrInput(strRifCtrl).ValoreCorrente
                  MskEd(Index).SetFocus
               End If
          End If
End With
End If
 '  On Local Error GoTo 0
Exit Sub
err_Imp:
    Call MXNU.MsgBoxEX(Err.Number & " " & Err.Description, vbCritical, 1007)
    On Local Error Resume Next
    MskEd(Index).SetFocus
    On Local Error GoTo 0
 
End Sub

Private Sub Scheda_Paint(Index As Integer)
    Call SchedaOmbreggiaControlli(Scheda(Index))
End Sub

Private Sub TxtBox_GotFocus(Index As Integer)
 SelContenuto TxtBox(Index)
End Sub

Private Sub TxtBox_KeyPress(Index As Integer, keyAscii As Integer)
Dim strRifCtrl As String
    With xAna
        strRifCtrl = .NomeControllo2NomeVariabile("TxtBox_" & Index)
        Call CtrlKey(keyAscii, xAna.GrInput(strRifCtrl).TipoInput)
    End With

End Sub

Private Sub TxtBox_LostFocus(Index As Integer)
Dim strRifCtrl As String
    'validazione + assegnamento campo
    strRifCtrl = xAna.NomeControllo2NomeVariabile("TxtBox_" & Index)
    
    If (xAna.GrInput(strRifCtrl).ValoreCorrente <> TxtBox(Index).text) Then
        If (strRifCtrl = xAna.NomeControlloCodice) Or Index = 7 Then
                Call xAna.AssegnaCampo(strRifCtrl, TxtBox(Index).text)
        End If
    End If
End Sub


Public Sub Svuota(TipoC As Variant)
    Dim i As Integer
    On Error Resume Next
    For i = 0 To UBound(TipoC) - 1
       TipoC(i) = ""
    Next i
End Sub

Public Property Let TipoOperazione(ByVal vData As String)
Dim i As Integer, k As Integer
Dim Valore As Variant
Dim strEntry As String
Dim cPar As CParametri

    mStrTipoOperazione = vData
    Set mColParametriOP = Nothing
    Set mColParametriOP = New Collection
    
    i = 1
    strEntry = MXNU.LeggiProfilo(MXNU.PercorsoPgm & "\MWSchedula.ini", mStrTipoOperazione, i, "")
    While Len(strEntry) > 0
        Valore = Split(strEntry, ";", , vbTextCompare)
        Set cPar = New CParametri
        With cPar
            .IDPar = i
            .NomePar = Valore(0)
            For k = 1 To UBound(Valore)
                .Tipo(k) = Valore(k)
                .Valore(k) = ""
            Next k
        End With
        mColParametriOP.Add cPar, Valore(0)
        'prossimo parametro
        i = i + 1
        strEntry = MXNU.LeggiProfilo(MXNU.PercorsoPgm & "\MWSchedula.ini", mStrTipoOperazione, i, "")
    Wend
        
End Property

Public Property Get TipoOperazione() As String
    TipoOperazione = mStrTipoOperazione
End Property

Public Property Get Parametro(ByVal vKeyPar As Variant) As CParametri
    On Local Error Resume Next
    Set Parametro = mColParametriOP(vKeyPar)
    On Local Error GoTo 0
End Property

Public Function ImpostaSchedulazione() As Boolean
    'mostro la form
    
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
    Me.Show vbModal
    'risultati
    Unload Me
    ImpostaSchedulazione = mBolSalvata
End Function


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

