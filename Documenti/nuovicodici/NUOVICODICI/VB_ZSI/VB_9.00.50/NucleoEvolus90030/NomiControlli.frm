VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0E5C731F-5D58-4240-BA4C-FB7B4C0BA48E}#1.0#0"; "mxctrl.ocx"
Begin VB.Form FrmNomiControlli 
   Caption         =   "Situazione Controlli"
   ClientHeight    =   5670
   ClientLeft      =   1620
   ClientTop       =   4020
   ClientWidth     =   11520
   Icon            =   "NomiControlli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11520
   Begin MXCtrl.MWSchedaBox MWSchedaBox1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   9975
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleWidth      =   11505
      ScaleHeight     =   5655
      Begin VB.CheckBox Chk 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Individuazione Controllo della Form"
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
         Height          =   315
         Left            =   5640
         TabIndex        =   7
         Top             =   5040
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Fine"
         Height          =   345
         Left            =   10050
         TabIndex        =   4
         Top             =   4920
         Width           =   1185
      End
      Begin VB.Frame FrmAna 
         Appearance      =   0  'Flat
         Caption         =   "Dettaglio Anagrafica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3465
         Left            =   5610
         TabIndex        =   3
         Top             =   600
         Width           =   5775
         Begin MXCtrl.MWSchedaBox SchFrameAna 
            Height          =   3135
            Left            =   60
            TabIndex        =   8
            Top             =   240
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   5530
            ForeColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LightColor      =   -2147483633
            ShadowColor     =   -2147483633
            ScaleWidth      =   5655
            ScaleHeight     =   3135
            FillWithGradient=   0   'False
            Begin MXCtrl.MWEtichetta etcro 
               Height          =   300
               Index           =   0
               Left            =   1620
               Top             =   60
               Width           =   3975
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
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   2
               Left            =   30
               Top             =   60
               Width           =   1515
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
               Caption         =   "Nome Campo DB"
            End
            Begin MXCtrl.MWEtichetta etcro 
               Height          =   300
               Index           =   1
               Left            =   1620
               Top             =   960
               Width           =   3975
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
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   0
               Left            =   30
               Top             =   960
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   476
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
               Caption         =   "Descrizione Et."
            End
            Begin MXCtrl.MWEtichetta etcro 
               Height          =   300
               Index           =   2
               Left            =   1620
               Top             =   1410
               Width           =   3975
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
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   1
               Left            =   30
               Top             =   1410
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   476
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
               Caption         =   "Valore Corrente"
            End
            Begin MXCtrl.MWEtichetta etcro 
               Height          =   300
               Index           =   3
               Left            =   1620
               Top             =   510
               Width           =   3975
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
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   3
               Left            =   30
               Top             =   510
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   476
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
               Caption         =   "Etichetta"
            End
            Begin MXCtrl.MWEtichetta etcro 
               Height          =   300
               Index           =   4
               Left            =   1620
               Top             =   2280
               Width           =   3975
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
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   4
               Left            =   30
               Top             =   2280
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   476
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
               Caption         =   "Nome Gruppo"
            End
            Begin MXCtrl.MWEtichetta etcro 
               Height          =   300
               Index           =   5
               Left            =   1620
               Top             =   1860
               Width           =   3975
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
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   5
               Left            =   30
               Top             =   1860
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   476
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
               Caption         =   "Validazione"
            End
            Begin MXCtrl.MWEtichetta etcro 
               Height          =   300
               Index           =   6
               Left            =   1620
               Top             =   2700
               Width           =   3975
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
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   6
               Left            =   30
               Top             =   2700
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   476
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
               Caption         =   "Formattazione"
            End
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
         Height          =   5475
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   30
         WhatsThisHelpID =   24024
         Width           =   5535
         Begin MSComctlLib.TreeView trwOggetti 
            Height          =   5175
            Left            =   120
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   180
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   9128
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
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":2832
               Key             =   "metodo98"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":2DD8
               Key             =   "entire"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":2EEA
               Key             =   "gruppo"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":3490
               Key             =   "utente"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":3A36
               Key             =   "formab"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":3B48
               Key             =   "formds"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":409A
               Key             =   "lingab"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":4194
               Key             =   "lingds"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":428E
               Key             =   "sitab"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":43A8
               Key             =   "sitds"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":44C2
               Key             =   "findab"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":4614
               Key             =   "findds"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":4766
               Key             =   "spreadab"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NomiControlli.frx":5428
               Key             =   "toolbuttonab"
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrmTab 
         Appearance      =   0  'Flat
         Caption         =   "Dettaglio Tabella"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         Left            =   5610
         TabIndex        =   5
         Top             =   630
         Visible         =   0   'False
         Width           =   5775
         Begin FPSpreadADO.fpSpread ssTab 
            Height          =   2955
            Left            =   150
            TabIndex        =   6
            Top             =   330
            Width           =   5505
            _Version        =   524288
            _ExtentX        =   9710
            _ExtentY        =   5212
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   2
            NoBeep          =   -1  'True
            SpreadDesigner  =   "NomiControlli.frx":5532
            AppearanceStyle =   0
         End
      End
      Begin MXCtrl.MWEtichetta etcro 
         Height          =   300
         Index           =   9
         Left            =   7290
         Top             =   4620
         Width           =   975
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
         Caption         =   "111"
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   9
         Left            =   5640
         Top             =   4620
         Width           =   1575
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
         Caption         =   "Help ID"
      End
      Begin MXCtrl.MWEtichetta etcro 
         Height          =   300
         Index           =   8
         Left            =   7290
         Top             =   4200
         Width           =   975
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
         Caption         =   "111"
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   8
         Left            =   5640
         Top             =   4200
         Width           =   1575
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
         Caption         =   "Num. Controlli"
      End
      Begin MXCtrl.MWEtichetta etcro 
         Height          =   300
         Index           =   7
         Left            =   7740
         Top             =   120
         Width           =   3375
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
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   7
         Left            =   5790
         Top             =   120
         Width           =   1875
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
         Caption         =   "Nome Controllo"
      End
   End
End
Attribute VB_Name = "FrmNomiControlli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1

Private Const PIC_FORM = "formab"
Private Const PIC_LING = "lingab"
Private Const PIC_CTRL = "formds"
Private Const PIC_SPREAD = "spreadab"
Private Const PIC_TOOLB = "toolbuttonab"

Public frmDef As Object
Public ListaCol As Collection

Dim MlngIDForm As Long
Dim MstrKeyForm As String
Dim MintNumCtrl As Integer
'rif.sch. A5408
Dim mbolLampeggia As Boolean

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    
    If frmDef Is Nothing Then GoTo FrmUnload
    
    MlngIDForm = frmDef.HelpContextID
    etcro(9).Caption = MlngIDForm
    MstrKeyForm = frmDef.NAME
    Call ssSpreadImposta(ssTab)
    '----------------------------
    'RIF.SCH.NR.1743
    'NON RISOLVIBILE IN ALTRA MANIERA SE NON GESTENDO
    'CASO PARTICOLARE PER FORM ORDINI DI PRODUZIONE
    '----------------------------
    If StrComp(MstrKeyForm, "frmordprod", vbTextCompare) = 0 Then
        Call TreeOrdProd_Inizializza
    ElseIf (StrComp(MstrKeyForm, "frmprogproduzione", vbTextCompare) = 0) Then
        'rif.A-4816 - gestione form progproduzione
        Call TreeProgProd_Inizializza
    ElseIf MstrKeyForm = "frmExtChild" Then
         MstrKeyForm = "objExt" 'frmDef.Controls(1).Name
         Call TreeEXT_Inizializza
    ElseIf (MstrKeyForm = "frmAnaArt") Or (MstrKeyForm = "frmAnaArtTip") Then
        'Anomalia nr. 5973
        Call TreeAnaArt_Inizializza
    ElseIf MstrKeyForm = "frmAnaConsGest" Then
        'Anomalia nr. 5742
        Call TreeCons_Inizializza
    ElseIf MstrKeyForm = "frmSchedaTrasporto" Then
        'Anomalia nr. 10487
        Call TreeSchTrasp_Inizializza
    Else
        Call TreeOggetti_Inizializza
    End If
    etcro(8).Caption = MintNumCtrl
    FrmTab.Visible = False
    FrmAna.Visible = True
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
    Call CentraFinestra(hwnd)
    Exit Sub
    
FrmUnload:
    Unload Me
    Exit Sub


ErrTrap:
Dim lngErrCod As Long
Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array(Me.Caption, lngErrCod, strErrDsc))

End Sub
Sub TreeEXT_Inizializza()
    Dim nodX As Node
    Dim ctrGen As Control
    Dim ctrParent As Object
    Dim strLingKey As String
    Dim bolEsisteLing As Boolean
    Dim vntDum As Variant
    Dim strKeyForm As String
    
    'imposto radice=form
    Set nodX = trwOggetti.Nodes.Add(, tvwFirst, MstrKeyForm, frmDef.Caption, PIC_FORM)
    nodX.Tag = CStr(ID_SCHEDA_FORM)
    Call nodX.EnsureVisible
                
    bolEsisteLing = False
    For Each ctrGen In frmDef.ContrEXT
        If TypeName(ctrGen) = "MWLinguetta" Then
            bolEsisteLing = True
            Exit For
        End If
    Next
    For Each ctrGen In frmDef.ContrEXT
        If Not bolEsisteLing Then
            If (TypeName(ctrGen) = "MWSchedaBox") Then
                    On Local Error Resume Next
                        If (ctrGen.Container.NAME = "") Then
                            Call TreeOggetti_GetChild(ctrGen, MstrKeyForm)
                        End If
                    On Local Error GoTo 0
            Else
                    On Local Error Resume Next
                        If (ctrGen.Container.NAME = "") Then
                            Call TreeOggetti_Add(ctrGen, MstrKeyForm)
                        End If
                    On Local Error GoTo 0
            End If
        Else
            If ControlloEXT(ctrGen, strKeyForm) Then
                'leggo i controlli contenuti nella form
                If TypeName(ctrGen) = "MWSchedaBox" Then
                    If strKeyForm <> "" Then
                        ' Prima Nota - schede Iva
                        Call TreeOggettiEXT_GetChildForm(ctrGen, strKeyForm)
                    Else
                        Call TreeOggettiEXT_GetChildForm(ctrGen, MstrKeyForm)
                    End If
                Else
                    On Local Error Resume Next
                    vntDum = trwOggetti.Nodes.Item(KeyControlloGet(ctrGen)).key
                    If Err.Number = 0 Then
                        Call TreeOggetti_Add(ctrGen, MstrKeyForm)
                    End If
                    On Local Error GoTo 0
                End If
            End If
        End If
    Next
    'imposto linguette
'    If frmDef.Controls(1).Name = "objExtWrapper" Then
 '       Set ctrParent = ContenitoreControlli(frmDef, True)
'    Else
 '       Set ctrParent = ContenitoreControlli(frmDef)
 '   End If
    
    Call TreeOggettiEXT_LingGet(frmDef, MstrKeyForm)
    
    Set nodX = Nothing
    Set ctrGen = Nothing
    Set ctrParent = Nothing

End Sub
Sub TreeOggettiEXT_LingGet(ctrParent As Object, ByVal strKeyParent As String)
    Dim nodX As Node
    Dim ctrGen As Control
    Dim strLingKey As String
    Dim schede As Object
    Dim ContrApp As Variant
    Dim boolParente As Boolean
    'On Local Error Resume Next
    
    If ctrParent.NAME = "frmExtChild" Then
        Set ContrApp = ctrParent.ContrEXT
        'Set schede = ctrParent.ContrEXT("Scheda")
    Else
        Set ContrApp = ctrParent.Controls
        'Set schede = ctrParent.Controls("Scheda")
    End If
   ' On Local Error GoTo 0
    
    'leggo le linguette contenute
    For Each ctrGen In ContrApp
         If (TypeName(ctrGen) = "MWLinguetta") Then
            On Error Resume Next
            boolParente = True
            
            boolParente = (ctrGen.Container.hwnd = ctrParent.hwnd)
            
            'Anomalia n.ro 7459
            'N.B. modifica necessaria a seguito della aggiunta della proprietà "hwnd" sulle estensioni
            '(necessaria per il corretto funzionamento dello zoom)
            ' eseguo il controllo della propreità name che andrà in errore nel caso il container sia lo user control
            If (ctrGen.Container.NAME <> TypeName(ctrGen.Container)) Then
                ' Il container non è uno user control
            Else
                boolParente = True
            End If
            ' Se vado in errore ctrGen.Container è lo user control estensivo
            If Err.Number <> 0 Then
                boolParente = True
                Err.Clear
            End If
            
            On Error GoTo 0
            If (boolParente And ctrGen.NAME <> "LingSit" And ctrGen.NAME <> "LingP" And ctrGen.NAME <> "LingEstrai") Then
                strLingKey = KeyControlloGet(ctrGen)
                Set nodX = trwOggetti.Nodes.Add(strKeyParent, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
                nodX.Tag = ctrGen.Index
                Call nodX.EnsureVisible
                'leggo i controlli contenuti nella linguetta
                If ctrParent.NAME = "frmExtChild" Then
                    Set schede = ctrParent.ContrEXT("Scheda", ctrGen.Index)
                    Call TreeOggettiEXT_GetChild(schede, strLingKey)
                    'cerca eventuali sottolinguette
                    Call TreeOggettiEXT_LingGet(schede, strLingKey)
                Else
                    Dim i As Integer, s As Integer, l As Integer, intIndice As Integer
                    l = Len(strLingKey)
                    s = InStr(strLingKey, "_")
                    intIndice = CInt(Mid(strLingKey, s + 1, (l - s)))
                    
                    For i = 0 To ctrParent.Controls.Count - 1
                        If TypeName(ctrParent.Controls(i)) = "MWSchedaBox" Then
                            If ctrParent.Controls(i).Index = intIndice Then
                                Set schede = ctrParent.Controls(i)
                                Call TreeOggettiEXT_GetChild(schede, strLingKey)
                                Call TreeOggettiEXT_LingGet(schede, strLingKey)
                                Exit For
                            End If
                        End If
                    Next i
                    
                    '** remmata a seguito correzione anomalia nr. 6096:
                    '** la seguente istruzione nel caso di una estensione non riesce a restituire
                    '** l'insieme di controlli scheda e dà l'errore "tipo non corrispondente"
                    'Set schede = ctrParent.Controls("Scheda")
                End If
                '** remmate a seguito correzione anomalia nr. 6096
'                Call TreeOggettiEXT_GetChild(schede, strLingKey)
'                'cerca eventuali sottolinguette
'                Call TreeOggettiEXT_LingGet(schede, strLingKey)
            End If
        End If
    Next
    Set nodX = Nothing
    Set ctrGen = Nothing
    Set schede = Nothing

End Sub
Sub TreeOggettiEXT_GetChild(ctrParent As Object, ByVal strKeyParent As String)
Dim nodX As Node
Dim ctrGen As Control
Dim BolContr As Boolean

    For Each ctrGen In frmDef.ContrEXT
        If (TypeName(ctrGen) <> "ImageList") And (TypeName(ctrGen) <> "CommonDialog") Then
            BolContr = False
            On Local Error Resume Next
            BolContr = (ctrGen.Container.hwnd = ctrParent.hwnd)
            On Local Error GoTo 0
            If BolContr Then
                If (TypeName(ctrGen) = "Frame") _
                    Or ((TypeName(ctrGen) = "MWSchedaBox") And (ctrGen.NAME <> "Scheda")) Then
                    Call TreeOggettiEXT_GetChild(ctrGen, strKeyParent)
                ElseIf (TypeName(ctrGen) <> "MWLinguetta") And (TypeName(ctrGen) <> "ImageList") And (TypeName(ctrGen) <> "MWSchedaBox") Then
                    Call TreeOggetti_Add(ctrGen, strKeyParent)
                End If
            End If
        End If
    Next
    Set ctrGen = Nothing

End Sub

Sub TreeOggettiEXT_GetChildForm(ctrParent As Object, ByVal strKeyParent As String)
Dim nodX As Node
Dim ctrGen As Control
Dim strLingKey As String
Dim intq As Integer
Dim bCtrl As Boolean

    For Each ctrGen In frmDef.ContrEXT
        If (TypeName(ctrGen) <> "ImageList") Then
            On Local Error Resume Next
            bCtrl = False
            bCtrl = (ctrGen.Container.hwnd = ctrParent.hwnd)
            On Local Error GoTo 0
            If bCtrl Then
                If (TypeName(ctrGen) = "Frame") Or ((TypeName(ctrGen) = "MWSchedaBox")) Then
                    'Call TreeOggetti_GetChildForm(ctrGen, strKeyParent)
                ElseIf (TypeName(ctrGen) = "MWLinguetta") Then
                    strLingKey = KeyControlloGet(ctrGen)
                    Set nodX = trwOggetti.Nodes.Add(strKeyParent, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
                    On Local Error Resume Next
                    nodX.Tag = ctrGen.Index
                    Call nodX.EnsureVisible
                    'leggo i controlli contenuti nella linguetta

                    
                    intq = frmDef.Scheda(ctrGen.Index).Index
                    If Err.Number = 0 Then
                        Call TreeOggettiEXT_GetChildForm(frmDef.Scheda(ctrGen.Index), strLingKey)
                    End If
                    
                    On Local Error GoTo 0
                ElseIf (TypeName(ctrGen) <> "ImageList") Then
                        Call TreeOggetti_Add(ctrGen, strKeyParent)
                        
                End If
            End If
        End If
    Next
    Set ctrGen = Nothing

End Sub

Sub TreeOggetti_LingGet(ctrParent As Object, ByVal strKeyParent As String)
    Dim nodX As Node
    Dim ctrGen As Control
    Dim strLingKey As String
    Dim schede As Object
    Dim schedeTP As Object
    Dim bolLingTP As Boolean

    On Local Error Resume Next
    Set schede = ContenitoreControlli(frmDef).Controls("Scheda")
    Set schedeTP = ContenitoreControlli(frmDef).Controls("SchedaTP")
    On Local Error GoTo ErrTrap
    
    'leggo le linguette contenute
    For Each ctrGen In ctrParent.Controls
        If (TypeName(ctrGen) = "MWLinguetta") Then
            'Rif. anmalia #7619 (commentata la riga successiva e ridotto il controllo dell'if - rimane il problema delle sottolinguette delle
            'linguette particolari tipo la linguetta delle situazioni
            'If (ctrGen.Container Is ctrParent And ctrGen.Name <> "LingSit" And ctrGen.Name <> "LingP" And ctrGen.Name <> "LingEstrai") Then
            If (ctrGen.Container Is ctrParent) Then
                strLingKey = KeyControlloGet(ctrGen)
                bolLingTP = (InStr(1, strLingKey, "LingTP", vbTextCompare) > 0)
                Set nodX = trwOggetti.Nodes.Add(strKeyParent, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
                'Rif. anomalia #7619
                If HasIndex(ctrGen) Then
                    nodX.Tag = ctrGen.Index
                End If
                Call nodX.EnsureVisible
                'Rif. anomalia #7619
                If HasIndex(ctrGen) Then
                    If Not bolLingTP Then
                        'leggo i controlli contenuti nella linguetta
                        If ExistObjArrayElement(frmDef.Scheda, ctrGen.Index) Then
                            Call TreeOggetti_GetChild(frmDef.Scheda(ctrGen.Index), strLingKey)
                            'cerca eventuali sottolinguette
                            Call TreeOggetti_LingGet(schede(ctrGen.Index), strLingKey)
                        End If
                    Else
                        'leggo i controlli contenuti nella linguetta
                        If ExistObjArrayElement(frmDef.SchedaTP, ctrGen.Index) Then
                            Call TreeOggetti_GetChild(frmDef.SchedaTP(ctrGen.Index), strLingKey)
                            'NB: problema con caricamento sottolinguette: le sottoschede non si chiamano "SchedaTP" e la ricorsione non funziona
                            'Call TreeOggetti_LingGet(schedeTP(ctrGen.Index), strLingKey)
                        End If
                    End If
                End If
            End If
        End If

ProssimoControllo:
    Next
    
    Set nodX = Nothing
    Set ctrGen = Nothing
    Set schede = Nothing
    
    Exit Sub
    
ErrTrap:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Errore in caricamento TreeView"
    Resume ProssimoControllo
    Resume
End Sub

Sub TreeOggetti_Add(ctrGen As Control, vntKeyParent As String)

    Dim nodX As Node
    
    On Error Resume Next
    
    MintNumCtrl = MintNumCtrl + 1
    Select Case TypeName(ctrGen)
        Case "MWEtichetta", "Label"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_CTRL)
        Case "MWLinguetta"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_LING)
        Case "TextBox"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_CTRL)
        Case "MaskEdBox"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_CTRL)
        Case "CommandButton"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_CTRL)
        Case "CheckBox"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_CTRL)
        Case "OptionButton"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_CTRL)
        Case "ComboBox"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_CTRL)
        Case "fpSpread"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_SPREAD)
        Case "ToolButton", "XPToolButton"
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_TOOLB)
        Case Else
            Set nodX = trwOggetti.Nodes.Add(vntKeyParent, tvwChild, KeyControlloGet(ctrGen), CtrlGetDsc(ctrGen), PIC_CTRL)
    End Select

    Set nodX = Nothing
    
    On Error GoTo 0
End Sub

Sub TreeOggetti_GetChildForm(ctrParent As Object, ByVal strKeyParent As String)
    Dim nodX As Node
    Dim ctrGen As Control
    Dim strLingKey As String
    Dim intq As Integer
    On Error GoTo TreeOggetti_ERR

    For Each ctrGen In frmDef
        If (TypeName(ctrGen) <> "ImageList") And (TypeName(ctrGen) <> "CommonDialog") And (TypeName(ctrGen) <> "DockingPane") And (TypeName(ctrGen) <> "CommandBars") Then
            strLingKey = TypeName(ctrGen)
            If (ctrGen.Container.hwnd = ctrParent.hwnd) Then
                If (TypeName(ctrGen) = "Frame") Or ((TypeName(ctrGen) = "MWSchedaBox")) Then
                    'Call TreeOggetti_GetChildForm(ctrGen, strKeyParent)
                ElseIf (TypeName(ctrGen) = "MWLinguetta") Then
                    strLingKey = KeyControlloGet(ctrGen)
                    If frmDef.NAME = "frmPnContabile" Then
                        On Local Error Resume Next
                    End If
                    Set nodX = trwOggetti.Nodes.Add(strKeyParent, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
                    If frmDef.NAME = "frmPnContabile" Then
                        If Err.Number <> 0 Then
                            GoTo TreeOggetti_END
                        End If
                    End If
                    nodX.Tag = ctrGen.Index
                    Call nodX.EnsureVisible
                    'leggo i controlli contenuti nella linguetta

                    On Local Error Resume Next
                    intq = frmDef.Scheda(ctrGen.Index).Index
                    If Err.Number = 0 Then
                        Call TreeOggetti_GetChildForm(frmDef.Scheda(ctrGen.Index), strLingKey)
                    End If
                    On Local Error GoTo 0
                ElseIf (TypeName(ctrGen) <> "ImageList") Then
                    Call TreeOggetti_Add(ctrGen, strKeyParent)
                End If
            End If
        End If
    Next
    Set ctrGen = Nothing

TreeOggetti_END:
    On Error GoTo 0
    Exit Sub

TreeOggetti_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("TreeOggetti", lngErrCod, strErrDsc))
    Resume TreeOggetti_END
    Resume
End Sub
Sub TreeOggetti_GetChild(ctrParent As Object, ByVal strKeyParent As String)
Dim nodX As Node
Dim ctrGen As Control

    For Each ctrGen In frmDef
        If (TypeName(ctrGen) <> "ImageList") And (TypeName(ctrGen) <> "CommonDialog") And (TypeName(ctrGen) <> "VT") And (TypeName(ctrGen) <> "Timer") And (TypeName(ctrGen) <> "DockingPane") And (TypeName(ctrGen) <> "CommandBars") Then
            If (ctrGen.Container.hwnd = ctrParent.hwnd) Then
                ' Rif. scheda # 8756 (aggiunto nell'if la PictureBox...)
                If (TypeName(ctrGen) = "Frame") _
                    Or (TypeName(ctrGen) = "PictureBox") Or ((TypeName(ctrGen) = "MWSchedaBox") And (ctrGen.NAME <> "Scheda")) Then
                    Call TreeOggetti_GetChild(ctrGen, strKeyParent)
                ElseIf (TypeName(ctrGen) <> "MWLinguetta") And (TypeName(ctrGen) <> "ImageList") And (TypeName(ctrGen) <> "MWSchedaBox") Then
                    Call TreeOggetti_Add(ctrGen, strKeyParent)
                End If
            End If
        End If
    Next
    Set ctrGen = Nothing

End Sub

Sub TreeAnaArt_Inizializza()
    Dim nodX As Node
    Dim ctrGen As Control
    Dim ctrParent As Object
    Dim strLingKey As String
    Dim vntDum As Variant
    Dim strKeyForm As String
    
    'imposto radice=form
    Set nodX = trwOggetti.Nodes.Add(, tvwFirst, MstrKeyForm, frmDef.Caption, PIC_FORM)
    nodX.Tag = CStr(ID_SCHEDA_FORM)
    Call nodX.EnsureVisible
                
    For Each ctrGen In frmDef.Controls
        'VT è il controllo per il Rational Visual Test
        If TypeName(ctrGen) <> "VT" Then
            If ControlloForm(ctrGen, strKeyForm) Then
                'leggo i controlli contenuti nella form
                If TypeName(ctrGen) = "MWSchedaBox" And ctrGen.NAME <> "SchedaD" Then
                    Call TreeOggetti_GetChildForm(ctrGen, MstrKeyForm)
                Else
                    On Local Error Resume Next
                    vntDum = trwOggetti.Nodes.Item(KeyControlloGet(ctrGen)).key
                    If Err.Number = 0 Then
                        Call TreeOggetti_Add(ctrGen, MstrKeyForm)
                    End If
                    On Local Error GoTo 0
                End If
            End If
        End If
    Next
    'Scheda duplica
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "SchedaD", Replace(MXNU.CaricaStringaRes(25024), "&", ""), "lingab")
    Call trwOggetti.Nodes.Add("SchedaD", tvwChild, "txtbD", "txtbD [" & MXNU.CaricaStringaRes(24116) & "]", PIC_CTRL)
    Call trwOggetti.Nodes.Add("SchedaD", tvwChild, "comD_1", "comD(1) [" & MXNU.CaricaStringaRes(25007) & "]", PIC_CTRL)
    Call trwOggetti.Nodes.Add("SchedaD", tvwChild, "comD_2", "comD(2) [" & MXNU.CaricaStringaRes(25008) & "]", PIC_CTRL)
    
    'imposto linguette
    Set ctrParent = ContenitoreControlli(frmDef)
    Call TreeOggetti_LingGet(ctrParent, MstrKeyForm)
    
    Set nodX = Nothing
    Set ctrGen = Nothing
    Set ctrParent = Nothing

End Sub


Sub TreeOggetti_Inizializza()
    Dim nodX As Node
    Dim ctrGen As Control
    Dim ctrParent As Object
    Dim strLingKey As String
    Dim bolEsisteLing As Boolean
    Dim vntDum As Variant
    Dim strKeyForm As String
    
    'imposto radice=form
    Set nodX = trwOggetti.Nodes.Add(, tvwFirst, MstrKeyForm, frmDef.Caption, PIC_FORM)
    nodX.Tag = CStr(ID_SCHEDA_FORM)
    Call nodX.EnsureVisible
                
    bolEsisteLing = False
    For Each ctrGen In frmDef.Controls
        If TypeName(ctrGen) = "MWLinguetta" Then
            bolEsisteLing = True
            Exit For
        End If
    Next
    For Each ctrGen In frmDef.Controls
        'VT è il controllo per il Rational Visual Test
        If TypeName(ctrGen) <> "VT" Then
            If Not bolEsisteLing Then
                If (TypeName(ctrGen) = "MWSchedaBox") Then
                    If (ctrGen.Container.hwnd = frmDef.hwnd) Then
                        Call TreeOggetti_GetChild(ctrGen, MstrKeyForm)
                    End If
                Else
                    On Error Resume Next
                    If (ctrGen.Container.hwnd = frmDef.hwnd) Then
                        Call TreeOggetti_Add(ctrGen, MstrKeyForm)
                    End If
                    On Error GoTo 0
                End If
            Else
                If ControlloForm(ctrGen, strKeyForm) Then
                    'leggo i controlli contenuti nella form
                    If TypeName(ctrGen) = "MWSchedaBox" Then
                        If strKeyForm <> "" Then
                            ' Prima Nota - schede Iva
                            Call TreeOggetti_GetChildForm(ctrGen, strKeyForm)
                        Else
                            Call TreeOggetti_GetChildForm(ctrGen, MstrKeyForm)
                        End If
                    Else
                        On Local Error Resume Next
                        vntDum = trwOggetti.Nodes.Item(KeyControlloGet(ctrGen)).key
                        If Err.Number = 0 Then
                            Call TreeOggetti_Add(ctrGen, MstrKeyForm)
                        End If
                        On Local Error GoTo 0
                    End If
                End If
            End If
        End If
    Next
    'imposto linguette
    Set ctrParent = ContenitoreControlli(frmDef)
    Call TreeOggetti_LingGet(ctrParent, MstrKeyForm)
    
    Set nodX = Nothing
    Set ctrGen = Nothing
    Set ctrParent = Nothing

End Sub
Private Function bolConsidera(ctrlGen As Control, flg As Boolean) As Boolean
    On Error Resume Next
    
    bolConsidera = False
    
    If (ctrlGen.Container.hwnd = frmDef.hwnd) Then
        bolConsidera = True
        flg = False
    Else
        If TypeName(ctrlGen) = "MWSchedaBox" Then
            If ctrlGen.NAME = "SchIva" Then
                bolConsidera = True
                flg = True
            End If
        End If
    End If
    
    On Error GoTo 0
End Function
Private Function bolConsideraEXT(ctrlGen As Control, flg As Boolean) As Boolean
    bolConsideraEXT = False
    On Local Error Resume Next
    
    If ctrlGen.Container.NAME = "" Then
        bolConsideraEXT = True
        flg = False
    Else
        If TypeName(ctrlGen) = "MWSchedaBox" Then
            If ctrlGen.NAME = "SchIva" Then
                bolConsideraEXT = True
                flg = True
            End If
        End If
    End If

End Function
Private Function ControlloForm(ctrlGen As Control, strKeyForm As String) As Boolean
    Dim bolResult As Boolean
    Dim intq As Integer
    Dim flg As Boolean

    bolResult = False
    strKeyForm = ""
    If TypeName(ctrlGen) = "ImageList" Or TypeName(ctrlGen) = "Timer" Then Exit Function
    If bolConsidera(ctrlGen, flg) Then
        Select Case TypeName(ctrlGen)
            Case "MWSchedaBox"
                If ctrlGen.NAME <> "SchedaSit" Then
                    If flg Then
                        bolResult = True
                        If ctrlGen.Index = 0 Then
                            strKeyForm = "LingIva_0"
                        Else
                            strKeyForm = "LingIva_1"
                        End If
                    Else
                        On Local Error Resume Next
                        intq = ctrlGen.Index
                        If Err.Number = 0 Then
                            intq = frmDef.LING(intq).Index
                            bolResult = (Err <> 0)
                        Else
                            bolResult = True
                        End If
                        On Local Error GoTo 0
                    End If
                End If
            Case "fpSpread"
                bolResult = True
        End Select
    End If
    ControlloForm = bolResult
    
End Function
Private Function ControlloEXT(ctrlGen As Control, strKeyForm As String) As Boolean
    Dim bolResult As Boolean
    Dim intq As Integer
    Dim flg As Boolean

    bolResult = False
    strKeyForm = ""
    If TypeName(ctrlGen) = "ImageList" Then Exit Function
    If bolConsideraEXT(ctrlGen, flg) Then
        Select Case TypeName(ctrlGen)
            Case "MWSchedaBox"
                If ctrlGen.NAME <> "SchedaSit" Then
                    If flg Then
                        bolResult = True
                        If ctrlGen.Index = 0 Then
                            strKeyForm = "LingIva_0"
                        Else
                            strKeyForm = "LingIva_1"
                        End If
                    Else
                        On Local Error Resume Next
                        intq = ctrlGen.Index
                        If Err.Number = 0 Then
                            Dim lingOBJ As Object
                            'Set lingOBJ = ContenitoreControlli(frmDef).Controls("ling")(intq)
                            Set lingOBJ = frmDef.ContrEXT("ling", intq)
                            'intq = frmDef.Ling(intq).Index
                            bolResult = (Err <> 0)
                        Else
                            bolResult = True
                        End If
                        On Local Error GoTo 0
                    End If
                End If
            Case "fpSpread"
                bolResult = True
        End Select
    End If
    ControlloEXT = bolResult


End Function
Function CtrlGetDsc(ctrlGen As Control) As String

    Dim vntIndex As Variant
    Dim strNom As String
    Dim strDsc As String
    Dim strKey As String
    Dim xAna As Object
    Dim strGruppo As String

    On Local Error Resume Next
    vntIndex = ctrlGen.Index
    'nome controllo
    If (Err = 0) Then
        strNom = ctrlGen.NAME & "(" & vntIndex & ")"
        strKey = ctrlGen.NAME & "_" & vntIndex
    Else
        strNom = ctrlGen.NAME
        strKey = strNom
    End If
    strDsc = ""
    For Each xAna In ListaCol
        If TypeName(xAna) = "Anagrafica" Then
            strGruppo = xAna.NomeControllo2NomeVariabile(strKey)
            If strGruppo <> "" Then
                strDsc = MXNU.CaricaCaptionInLingua(xAna.grinput(strGruppo).DescrizioneEtichetta)
            End If
        End If
    Next
    On Local Error GoTo 0
    strNom = strNom & " " & strDsc
    Select Case TypeName(ctrlGen)
        Case "TextBox"
            If ctrlGen.text <> "" Then strNom = strNom & " [" & ctrlGen.text & "]"
        Case "Label", "MWLinguetta", "CommandButton", "CheckBox", "OptionButton" ' rif. scheda #8756 (aggiunto gli option button)
            If ctrlGen.Caption <> "" Then strNom = strNom & " [" & ctrlGen.Caption & "]"
    End Select
    CtrlGetDsc = strNom
End Function

Private Sub Form_Paint()
    Call SchedaOmbreggiaControlli(Me)
End Sub

'rif.sch. A5408
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mbolLampeggia Then
        Cancel = -1
    End If
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
    Set frmDef = Nothing
    Set ListaCol = Nothing
    Set FrmNomiControlli = Nothing
End Sub

Private Sub SchFrameAna_Paint()
    SchedaOmbreggiaControlli SchFrameAna
End Sub


Private Sub trwOggetti_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim ctrlGen As Control
    Dim ctrlLing As Control
    Dim t@, i%
    Dim xAna As Object
    Dim strGruppo As String
    Dim bolTrovato As Boolean
    Dim vntDes As Variant
    Dim bolVisible  As Boolean
    Dim BoolESTENSIONE As Boolean
    
    BoolESTENSIONE = (frmDef.NAME = "frmExtChild")
    If MXVA.TrovatoControlloForm(frmDef, Node.key, ctrlGen, , , BoolESTENSIONE) Then
        ' Le righe documenti vanno in loop
        If Not (ctrlGen.NAME = "ssRighe" And frmDef.HelpContextID = 4000) Then
            bolTrovato = False
            ssTab.Row = -1
            ssTab.Col = -1
            ssTab.Action = ActionClearText
            etcro(7) = Node.key
            FrmTab.Visible = False
            FrmAna.Visible = True
            DoEvents
            If Not (ListaCol Is Nothing) Then
                For Each xAna In ListaCol
                    If TypeName(xAna) = "Anagrafica" Then
                        strGruppo = xAna.NomeControllo2NomeVariabile(Node.key)
                        If strGruppo <> "" Then
                            etcro(0).Caption = xAna.grinput(strGruppo).NomeCampoDb
                            etcro(3).Caption = xAna.grinput(strGruppo).ControlloEtichetta
                            etcro(1).Caption = xAna.grinput(strGruppo).DescrizioneEtichetta & " - " & MXNU.CaricaCaptionInLingua(xAna.grinput(strGruppo).DescrizioneEtichetta)
                            etcro(2).Caption = xAna.grinput(strGruppo).ValoreCorrente
                            etcro(5).Caption = xAna.grinput(strGruppo).TipoValidazione
                            etcro(4).Caption = xAna.grinput(strGruppo).key
                            etcro(6).Caption = xAna.grinput(strGruppo).Formattazione
                            frmDef.Refresh
                            DoEvents
                            bolTrovato = True
                            Exit For
                        End If
                    ElseIf TypeName(xAna) = "CTabelle" And TypeName(ctrlGen) = "fpSpread" Then
                        If xAna.NOMEFOGLIO = MXNU.nomeControllo(ctrlGen) Then
                            ssTab.MaxRows = 500
                            FrmTab.Visible = True
                            FrmAna.Visible = False
                            For i = 1 To xAna.ColonneEffettive
                                Call ssTab.SetText(2, i, xAna.NomeCampo(CStr(i)))
                                vntDes = ssCellGetValue(ctrlGen, i, 0)
                                Call ssTab.SetText(1, i, CStr(vntDes))
                            Next i
                            ssTab.MaxRows = i - 1
                            For i = 0 To 6: etcro(i).Caption = "": Next i
                            etcro(7) = Node.key
                            Exit For
                        End If
                    End If
                Next
            End If
            If Not bolTrovato Then
                For i = 0 To 6: etcro(i).Caption = "": Next i
                etcro(1).ToolTipText = ""
                'Sviluppo nr. 1569
                On Local Error Resume Next
                If TypeName(ctrlGen) = "MWEtichetta" Or TypeName(ctrlGen) = "Label" Then
                    bolTrovato = False
                    If Not (ListaCol Is Nothing) Then
                        For Each xAna In ListaCol
                            If TypeName(xAna) = "Anagrafica" Then
                                For i = 1 To xAna.grinput.Count
                                    If StrComp(Node.key, xAna.grinput(i).ControlloEtichetta, vbTextCompare) = 0 Then
                                        etcro(1).Caption = xAna.grinput(i).DescrizioneEtichetta & " - " & MXNU.CaricaCaptionInLingua(xAna.grinput(i).DescrizioneEtichetta)
                                        bolTrovato = True
                                        Exit For
                                    End If
                                Next i
                            End If
                            If bolTrovato Then Exit For
                        Next
                    End If
                End If
                If Not bolTrovato And ctrlGen.WhatsThisHelpID <> 0 Then
                    etcro(1).Caption = "{" & ctrlGen.WhatsThisHelpID & "} - " & ctrlGen.Caption
                End If
                On Local Error GoTo 0
            End If
            DoEvents
            On Local Error Resume Next
            If Chk.value = vbChecked Then
                frmDef.ZOrder 0
                DoEvents
                If MXVA.TrovatoControlloForm(frmDef, Node.Parent.key, ctrlLing) Then
                    Call ctrlLing.SetFocus
                    DoEvents
                End If
                On Local Error Resume Next
                bolVisible = ctrlGen.Container.Visible
                If bolVisible And Err = 0 Then
                    bolVisible = ctrlGen.Visible
                    mbolLampeggia = True ' rif.sch. A5408
                    For i = 1 To 5
                        ctrlGen.Visible = False
                        frmDef.Refresh
                        DoEvents
                        t = Timer
                        While (Timer - t) <= 0.5:: Wend
                        ctrlGen.Visible = True
                        frmDef.Refresh
                        DoEvents
                        t = Timer
                        While (Timer - t) <= 0.5:: Wend
                    Next i
                    mbolLampeggia = False ' rif.sch. A5408
                    ctrlGen.Visible = bolVisible
                End If
            End If
            Set ctrlLing = Nothing
            Me.ZOrder 0
            On Local Error GoTo 0
        End If
    End If
    Set xAna = Nothing
    Set ctrlGen = Nothing
    
End Sub

Private Sub TreeOrdProd_Inizializza()
Dim nodX As Node
Dim strLingKey As String
Dim strLingKeyRighe As String
Dim ctrGen As Object
Const SCHEDA_TESTA = 0
Const SCHEDA_RIGHE = 1
Const SCHEDA_PIEDE = 2
Const SCHEDA_ORDINI = 3
Const SCHEDA_IMPEGNI = 4
Const SCHEDA_LAVFASE = 5

    On Local Error GoTo err_TreeOrdProd
    'imposto radice=form
    Set nodX = trwOggetti.Nodes.Add(, tvwFirst, MstrKeyForm, frmDef.Caption, PIC_FORM)
    nodX.Tag = CStr(ID_SCHEDA_FORM)
    Call nodX.EnsureVisible
    Set nodX = Nothing

    'linguetta testa
    Set ctrGen = frmDef.LING(SCHEDA_TESTA)
    strLingKey = KeyControlloGet(ctrGen)
    Set nodX = trwOggetti.Nodes.Add(MstrKeyForm, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
    nodX.Tag = ctrGen.Index
    Call nodX.EnsureVisible
    Set nodX = Nothing
    Call TreeOggetti_GetChild(frmDef.Scheda(SCHEDA_TESTA), strLingKey)
    
    'linguetta righe
    Set ctrGen = frmDef.LING(SCHEDA_RIGHE)
    strLingKeyRighe = KeyControlloGet(ctrGen)
    Set nodX = trwOggetti.Nodes.Add(MstrKeyForm, tvwChild, strLingKeyRighe, swapp(ctrGen.Caption, "&", ""), "lingab")
    nodX.Tag = ctrGen.Index
    Call nodX.EnsureVisible
    Set nodX = Nothing
    
    'linguetta piede
    Set ctrGen = frmDef.LING(SCHEDA_PIEDE)
    strLingKey = KeyControlloGet(ctrGen)
    Set nodX = trwOggetti.Nodes.Add(MstrKeyForm, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
    nodX.Tag = ctrGen.Index
    Call nodX.EnsureVisible
    Call TreeOggetti_GetChild(frmDef.Scheda(SCHEDA_PIEDE), strLingKey)
    Set nodX = Nothing
    
    'linguetta ordini
    Set ctrGen = frmDef.LING(SCHEDA_ORDINI)
    strLingKey = KeyControlloGet(ctrGen)
    Set nodX = trwOggetti.Nodes.Add(strLingKeyRighe, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
    nodX.Tag = ctrGen.Index
    Call nodX.EnsureVisible
    Call TreeOggetti_GetChild(frmDef.Scheda(SCHEDA_ORDINI), strLingKey)
    Set nodX = Nothing
    
    'linguetta impegni
    Set ctrGen = frmDef.LING(SCHEDA_IMPEGNI)
    strLingKey = KeyControlloGet(ctrGen)
    Set nodX = trwOggetti.Nodes.Add(strLingKeyRighe, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
    nodX.Tag = ctrGen.Index
    Call nodX.EnsureVisible
    Call TreeOggetti_GetChild(frmDef.Scheda(SCHEDA_IMPEGNI), strLingKey)
    Set nodX = Nothing
    
    'RIF.A#9392 - linguetta lavorazioni di fase
    Set ctrGen = frmDef.LING(SCHEDA_LAVFASE)
    strLingKey = KeyControlloGet(ctrGen)
    Set nodX = trwOggetti.Nodes.Add(strLingKeyRighe, tvwChild, strLingKey, swapp(ctrGen.Caption, "&", ""), "lingab")
    nodX.Tag = ctrGen.Index
    Call nodX.EnsureVisible
    Call TreeOggetti_GetChild(frmDef.SchedaBack(0), strLingKey) 'RIF.A#9392
    Call TreeOggetti_GetChild(frmDef.Scheda(SCHEDA_LAVFASE), strLingKey)
    Set nodX = Nothing
    
fine_TreeOrdProd:
    Set ctrGen = Nothing
    Set nodX = Nothing
    On Local Error GoTo 0
    Exit Sub
    
err_TreeOrdProd:
    MXNU.MsgBoxEX 1009, vbCritical, 1007, Array("Inizializza Oggetti", Err.Number, Err.Description)
    Resume fine_TreeOrdProd
End Sub


'Riferimento Scheda Anomalia nr. 2717
'Nel caso esistano linguette che non hanno una scheda associata non vengono fatti test per trovare sottoschede o altro.
Private Function ExistObjArrayElement(obj As Object, Index As Integer) As Boolean
    Dim strNome As String
    On Error Resume Next
    
    strNome = obj(Index).Index
    
    If Err <> 0 Then
        ExistObjArrayElement = False
    Else
        ExistObjArrayElement = True
    End If
    
    On Error GoTo 0
End Function

'rif.A#4816 - gestione personalizzata per la form progproduzione
Private Sub TreeProgProd_Inizializza()
Dim ctrGen As Control
Dim nodX As MSComctlLib.Node

    'imposto radice=form
    Set nodX = trwOggetti.Nodes.Add(, tvwFirst, MstrKeyForm, frmDef.Caption, PIC_FORM)
    nodX.Tag = CStr(ID_SCHEDA_FORM)
    Call nodX.EnsureVisible
    'imposto figli = controlli
    For Each ctrGen In frmDef.Controls
        If (TypeName(ctrGen) = "MWSchedaBox") Then
            If (ctrGen.Container.hwnd = frmDef.hwnd) Then
                Call TreeOggetti_GetChild(ctrGen, MstrKeyForm)
            End If
        Else
            On Error Resume Next
            If (ctrGen.Container.hwnd = frmDef.hwnd) Then
                Call TreeOggetti_Add(ctrGen, MstrKeyForm)
            End If
            On Error GoTo 0
        End If
    Next ctrGen
End Sub

Private Sub TreeSchTrasp_Inizializza()
    Dim nodX As Node
    Dim ctrGen As Control
    Dim ctrParent As Object
    Dim strLingKey As String
    Dim vntDum As Variant
    Dim strKeyForm As String
    Dim i As Integer
    
    On Local Error GoTo err_TreeSchTrasp
    'La form è fuori standard
    Set nodX = trwOggetti.Nodes.Add(, tvwFirst, MstrKeyForm, frmDef.Caption, PIC_FORM)
    nodX.Tag = CStr(ID_SCHEDA_FORM)
    Call nodX.EnsureVisible
                
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "txtb_0", "txtb(0)[" & frmDef.Txtb(0).text & "]", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "tbSel_0", "tbSel(0)", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "tbSel_1", "tbSel(1)", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "sssped", "sssped", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "ssTrRighe", "ssTrRighe", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "txtb_1", "txtb(1)[" & frmDef.Txtb(1).text & "]", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "cmd_0", "cmd_0[" & Replace(frmDef.Cmd(0).Caption, "&", "") & "]", PIC_CTRL)
    For i = 2 To 10
        Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "txtb_" & i, "txtb(" & i & ")[" & CtrlGetDsc(frmDef.Txtb(i)) & "]", PIC_CTRL)
    Next i
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "mskEdt_0", "mskEdt_0[" & frmDef.MskEdt(0).text & "]", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "cmd_1", "cmd_1[" & Replace(frmDef.Cmd(1).Caption, "&", "") & "]", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "cmd_2", "cmd_2[" & Replace(frmDef.Cmd(2).Caption, "&", "") & "]", PIC_CTRL)
    
fine_TreeSchTrasp:
    On Local Error GoTo 0
    Set nodX = Nothing
    Set ctrGen = Nothing
    Set ctrParent = Nothing
    Exit Sub
    
err_TreeSchTrasp:
    MXNU.MsgBoxEX 1009, vbCritical, 1007, Array("Inizializza Oggetti", Err.Number, Err.Description)
    Resume fine_TreeSchTrasp

End Sub

Private Sub TreeCons_Inizializza()
    Dim nodX As Node
    Dim ctrGen As Control
    Dim ctrParent As Object
    Dim strLingKey As String
    Dim vntDum As Variant
    Dim strKeyForm As String
    
    'La form di analisi consegne è completamente fuori standard, anche come nomi schede
    'quindi non rimane che fare questo codice
    'imposto radice=form
    Set nodX = trwOggetti.Nodes.Add(, tvwFirst, MstrKeyForm, frmDef.Caption, PIC_FORM)
    nodX.Tag = CStr(ID_SCHEDA_FORM)
    Call nodX.EnsureVisible
                
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "optB_0", "optB(0) [" & frmDef.optB(0).Caption & "]", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "optB_1", "optB(1) [" & frmDef.optB(1).Caption & "]", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "optB_2", "optB(2) [" & frmDef.optB(2).Caption & "]", PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "etc_0", CtrlGetDsc(frmDef.etc(0)), PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "cmb_0", "cmb_0", PIC_CTRL)
    
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "cmdAnnulla_0", CtrlGetDsc(frmDef.cmdAnnulla(0)), PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "cmdAv_0", CtrlGetDsc(frmDef.cmdAv(0)), PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "cmdAv_1", CtrlGetDsc(frmDef.cmdAv(1)), PIC_CTRL)
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "cmdInd_0", CtrlGetDsc(frmDef.cmdInd(0)), PIC_CTRL)
                
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "Ling_0", swapp(frmDef.LING(0).Caption, "&", ""), "lingab")
    Call trwOggetti.Nodes.Add("Ling_0", tvwChild, "ctlFiltro(1)", "ctlFiltro(1)", PIC_CTRL)
    Call trwOggetti.Nodes.Add("Ling_0", tvwChild, "ssFiltroDoc", "ssFiltroDoc", PIC_CTRL)
    Call trwOggetti.Nodes.Add("Ling_0", tvwChild, "ssFiltroMag", "ssFiltroMag", PIC_CTRL)
    Call trwOggetti.Nodes.Add("Ling_0", tvwChild, "ssPenna", "ssPenna", PIC_CTRL)
    Call trwOggetti.Nodes.Add("Ling_0", tvwChild, "comPenna_0", CtrlGetDsc(frmDef.comPenna(0)), PIC_CTRL)
    Call trwOggetti.Nodes.Add("Ling_0", tvwChild, "comPenna_1", CtrlGetDsc(frmDef.comPenna(1)), PIC_CTRL)
                
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "Ling_1", swapp(frmDef.LING(1).Caption, "&", ""), "lingab")
    Call trwOggetti.Nodes.Add("Ling_1", tvwChild, "ssFiltroDaEv", "ssFiltroDaEv", PIC_CTRL)
    Call trwOggetti.Nodes.Add("Ling_1", tvwChild, "ctlFiltro(2)", "ctlFiltro(2)", PIC_CTRL)
                
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "Ling2_0", swapp(frmDef.ling2(0).Caption, "&", ""), "lingab")
    Call trwOggetti.Nodes.Add("Ling2_0", tvwChild, "etc_2", CtrlGetDsc(frmDef.etc(2)), PIC_CTRL)
    Call trwOggetti.Nodes.Add("Ling2_0", tvwChild, "txtb_0", "txtb_0", PIC_CTRL)
    Call trwOggetti.Nodes.Add("Ling2_0", tvwChild, "ssRisult", "ssRisult", PIC_CTRL)
                
    Call trwOggetti.Nodes.Add(nodX.key, tvwChild, "Ling2_1", swapp(frmDef.ling2(1).Caption, "&", ""), "lingab")
    Call trwOggetti.Nodes.Add("Ling2_1", tvwChild, "ssTotali", "ssTotali", PIC_CTRL)
    
fine_TreeCons:
    Set nodX = Nothing
    Set ctrGen = Nothing
    Set ctrParent = Nothing
    Exit Sub
    
err_TreeCons:
    MXNU.MsgBoxEX 1009, vbCritical, 1007, Array("Inizializza Oggetti", Err.Number, Err.Description)
    Resume fine_TreeCons
End Sub





'Per Metodo Evolus
Private Sub mResize_AfterResize()
    Call AvvicinaLing(Me)
End Sub

