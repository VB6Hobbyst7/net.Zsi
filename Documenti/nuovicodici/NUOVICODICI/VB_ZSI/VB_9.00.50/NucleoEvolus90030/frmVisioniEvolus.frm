VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CODEJOCK.COMMANDBARS.V15.3.1.OCX"
Object = "{90F671A1-7F84-49E0-9480-79F8C459AD2A}#1.0#0"; "MXKit.ocx"
Object = "{6764463D-875B-4A07-905E-B847D469D061}#1.0#0"; "mxctrl.ocx"
Begin VB.Form frmVisioni 
   Appearance      =   0  'Flat
   ClientHeight    =   7515
   ClientLeft      =   4830
   ClientTop       =   3360
   ClientWidth     =   11385
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
   Icon            =   "frmVisioniEvolus.frx":0000
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7515
   ScaleWidth      =   11385
   Begin FPSpreadADO.fpSpread SpreadConf 
      Height          =   195
      Left            =   9060
      TabIndex        =   28
      Top             =   60
      Visible         =   0   'False
      Width           =   675
      _Version        =   524288
      _ExtentX        =   1191
      _ExtentY        =   344
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      SpreadDesigner  =   "frmVisioniEvolus.frx":0442
   End
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
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
      Caption         =   "&Filtro"
   End
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      WhatsThisHelpID =   21056
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
      Caption         =   "&Visione"
   End
   Begin MXCtrl.MWLinguetta Ling 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      WhatsThisHelpID =   21057
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
      Caption         =   "&Totali"
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6300
      Index           =   1
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   11113
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
      ScaleWidth      =   11355
      ScaleHeight     =   6300
      Begin MXCtrl.MWSchedaBox SchedaTrovaBox 
         Height          =   1155
         Left            =   30
         TabIndex        =   19
         Top             =   5145
         Visible         =   0   'False
         Width           =   11325
         _ExtentX        =   19976
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
         ScaleWidth      =   11325
         ScaleHeight     =   1155
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
            TabIndex        =   22
            Top             =   30
            Width           =   7000
            Begin VB.CommandButton ComEsciTrova 
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   6720
               TabIndex        =   32
               Top             =   0
               Width           =   255
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
               TabIndex        =   25
               Top             =   300
               Visible         =   0   'False
               Width           =   5115
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
               TabIndex        =   24
               Top             =   300
               Width           =   5115
            End
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
               TabIndex        =   23
               Top             =   0
               Visible         =   0   'False
               Width           =   1935
            End
            Begin FPSpreadADO.fpSpread ssTrova 
               Height          =   975
               Left            =   120
               TabIndex        =   26
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
               SpreadDesigner  =   "frmVisioniEvolus.frx":0A29
               UserResize      =   0
               VisibleCols     =   6
               VisibleRows     =   3
               AppearanceStyle =   0
            End
            Begin MXCtrl.MWSchedaBox SchBottTrova 
               Height          =   735
               Left            =   5340
               TabIndex        =   29
               Top             =   200
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   1296
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Bevel           =   0
               LightColor      =   6908265
               BevelWidth      =   0
               ScaleWidth      =   1515
               ScaleHeight     =   735
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
                  Height          =   315
                  Left            =   60
                  TabIndex        =   31
                  Top             =   75
                  Width           =   1410
               End
               Begin VB.CommandButton comAvanzato 
                  Caption         =   "Avanzate >>"
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
                  Left            =   60
                  TabIndex        =   30
                  Top             =   435
                  Width           =   1410
               End
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
            Height          =   975
            Index           =   2
            Left            =   6780
            TabIndex        =   20
            Top             =   30
            Visible         =   0   'False
            Width           =   4335
            Begin FPSpreadADO.fpSpread ssOpzioni 
               Height          =   735
               Index           =   0
               Left            =   180
               TabIndex        =   21
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
               SpreadDesigner  =   "frmVisioniEvolus.frx":2257
               VisibleCols     =   1
               VisibleRows     =   3
               AppearanceStyle =   0
            End
         End
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
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
            WhatsThisHelpID =   24015
            Width           =   4335
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
         Height          =   6135
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   11055
         Begin FPSpreadADO.fpSpread ssVisione 
            DragIcon        =   "frmVisioniEvolus.frx":31E8
            Height          =   5760
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   10815
            _Version        =   524288
            _ExtentX        =   19076
            _ExtentY        =   10160
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
            SpreadDesigner  =   "frmVisioniEvolus.frx":34F2
            UnitType        =   2
            UserResize      =   0
            VirtualOverlap  =   15
            VirtualRows     =   15
            VisibleCols     =   10
            VisibleRows     =   16
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   1
            AppearanceStyle =   0
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
         Index           =   0
         Left            =   7020
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   60
         Visible         =   0   'False
         Width           =   4035
      End
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6300
      Index           =   2
      Left            =   0
      TabIndex        =   14
      Top             =   1200
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   11113
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
      ScaleWidth      =   11355
      ScaleHeight     =   6300
      Begin MXCtrl.MWLinguetta Ling 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   120
         WhatsThisHelpID =   21058
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
         Caption         =   "&Valori"
      End
      Begin MXCtrl.MWLinguetta Ling 
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   12
         Top             =   120
         WhatsThisHelpID =   21059
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
         Caption         =   "&Grafico"
      End
      Begin MXCtrl.MWSchedaBox Scheda 
         Height          =   5655
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   11055
         _ExtentX        =   19500
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
         LightColor      =   6908265
         ScaleWidth      =   11055
         ScaleHeight     =   5655
         Begin VB.Frame Frame 
            Appearance      =   0  'Flat
            Caption         =   "Totalizzazione Corrente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   60
            WhatsThisHelpID =   24016
            Width           =   5475
            Begin VB.ComboBox cmbTotali 
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
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   240
               Width           =   5235
            End
         End
         Begin FPSpreadADO.fpSpread ssTotali 
            Height          =   4620
            Left            =   120
            TabIndex        =   11
            Top             =   900
            Width           =   10755
            _Version        =   524288
            _ExtentX        =   18971
            _ExtentY        =   8149
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
               Name            =   "Tahoma"
               Size            =   9
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
            SpreadDesigner  =   "frmVisioniEvolus.frx":39F3
            UnitType        =   2
            UserResize      =   0
            VirtualOverlap  =   15
            VirtualRows     =   15
            VisibleCols     =   10
            VisibleRows     =   16
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   1
            AppearanceStyle =   0
         End
      End
      Begin MXCtrl.MWSchedaBox Scheda 
         Height          =   5655
         Index           =   4
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   11055
         _ExtentX        =   19500
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
         LightColor      =   6908265
         ScaleWidth      =   11055
         ScaleHeight     =   5655
         Begin MXKit.CTLGrafico objChart 
            Height          =   5535
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   10755
            _ExtentX        =   18971
            _ExtentY        =   9763
         End
      End
   End
   Begin MXCtrl.MWSchedaBox Scheda 
      Height          =   6300
      Index           =   0
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   11113
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
      ScaleWidth      =   11355
      ScaleHeight     =   6300
      Begin MXKit.ctlImpostazioni ctlImpFiltro 
         Height          =   495
         Left            =   1665
         TabIndex        =   16
         Top             =   120
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   873
      End
      Begin FPSpreadADO.fpSpread ssFiltroDati 
         Height          =   5295
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   10995
         _Version        =   524288
         _ExtentX        =   19394
         _ExtentY        =   9340
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         SpreadDesigner  =   "frmVisioniEvolus.frx":3EF4
         AppearanceStyle =   0
      End
   End
   Begin MSComctlLib.ImageList ImgListTB 
      Left            =   6300
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":4343
            Key             =   "Crystal"
            Object.Tag             =   "2220"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":5397
            Key             =   "Excel"
            Object.Tag             =   "2200"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":63EB
            Key             =   "Indietro"
            Object.Tag             =   "2104"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":743F
            Key             =   "Seleziona"
            Object.Tag             =   "2105"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":8493
            Key             =   "Primo"
            Object.Tag             =   "2100"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":94E7
            Key             =   "Precedente"
            Object.Tag             =   "2101"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":A53B
            Key             =   "Ultimo"
            Object.Tag             =   "2103"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":B58F
            Key             =   "Successivo"
            Object.Tag             =   "2102"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":C5E3
            Key             =   "Termina"
            Object.Tag             =   "2106"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":D637
            Key             =   "Ricarica"
            Object.Tag             =   "2107"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":E68B
            Key             =   "MOLAP"
            Object.Tag             =   "2240"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":F6DF
            Key             =   "OOCalc"
            Object.Tag             =   "2210"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":10733
            Key             =   "Info"
            Object.Tag             =   "2110"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":11787
            Key             =   "Filtro"
            Object.Tag             =   "2108"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":127DB
            Key             =   "OldVisione"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":1382F
            Key             =   "Raggruppa"
            Object.Tag             =   "2150"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":14109
            Key             =   "RaggruppaChiudi"
            Object.Tag             =   "2250"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":1495B
            Key             =   "Legenda"
            Object.Tag             =   "2109"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":159AF
            Key             =   "Azioni"
            Object.Tag             =   "2170"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":16A03
            Key             =   "LingVisione"
            Object.Tag             =   "2180"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":17655
            Key             =   "LingTotali"
            Object.Tag             =   "2190"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":182A7
            Key             =   "RibbonMaximize"
            Object.Tag             =   "2301"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":185F9
            Key             =   "RibbonMinimize"
            Object.Tag             =   "2302"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":1894B
            Key             =   "StampaTotali"
            Object.Tag             =   "2310"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":1959D
            Key             =   "OpzioniVisione"
            Object.Tag             =   "2320"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":1A1EF
            Key             =   "Trova"
            Object.Tag             =   "2350"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":1AE41
            Key             =   "Visione"
            Object.Tag             =   "2120"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":1BA93
            Key             =   "TotaliRefresh"
            Object.Tag             =   "2353"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisioniEvolus.frx":1C6E5
            Key             =   "QLikView"
            Object.Tag             =   "2270"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   5640
      Top             =   120
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmVisioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'Per Metodo Evolus
Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1
'Attribute mResize.VB_VarHelpID = -1

Public FormProp As New CFormProp

Private Enum enmTipoSaldiIniziali
    enmSommaSaldiAP = 0
    enmMovApertura = 1
    enmMovChiusuraAP = 2
End Enum
Private MenmTipoSaldi As enmTipoSaldiIniziali

'======================================================
'           definizione costanti
'======================================================
Const SCH_FILTRO = 0
Const SCH_VISIONE = 1
Const SCH_TOTALI = 2
Const SSCH2_VALORI = 3
Const SSCH2_GRAFICO = 4


'Ribbon Visione
Const ID_TABVIS_DATI = 0
Const ID_TABVIS_EXPORT = 1

Const ID_VISGROUP_GROUPSEZ = 1
Const ID_VISGROUP_GROUPVIS = 2
Const ID_VISGROUP_GROUPAZ = 3
Const ID_VISGROUP_GROUPINFO = 4
Const ID_VISGROUP_GROUPEXP = 5
Const ID_VISGROUP_GROUPTOTALI = 6

Const ID_VIS_FILTRO = 2108
Const ID_VIS_VISIONE = 2180
Const ID_VIS_TOTALI = 2190
Const ID_VIS_TIPOVIS = 2120
Const ID_VIS_RAGGRUPPA = 2150
Const ID_VIS_RAGGRUPPACHIUDI = 2250
Const ID_VIS_AZIONI = 2170
Const ID_VIS_EXPORTEXCEL = 2200
Const ID_VIS_EXPORTOOCALC = 2210
Const ID_VIS_EXPORTCRYSTAL = 2220
Const ID_VIS_EXPORTOLAP = 2240
Const ID_VIS_EXPORTQLIK = 2270
Const ID_VIS_RICARICA = 2107
Const ID_VIS_LEGENDA = 2109
Const ID_VIS_INFO = 2110
Const ID_VIS_OPZIONIVIS = 2320
Const ID_VIS_STAMPARAGGR = 2310
Const ID_VIS_TROVA = 2350
Const ID_VIS_TROVASTD = 2351
Const ID_VIS_TROVAAVZ = 2352
Const ID_VIS_RICTOTALI = 2353
Const ID_RIBBON_MINIMIZE = 2302
Const ID_RIBBON_EXPAND = 2301

Const IMAGEBASE = 10000

'======================================================
'           definizione classi visioni
'======================================================
Private WithEvents cTraccia As MXKit.cTraccia
Attribute cTraccia.VB_VarHelpID = -1
Private cInterfaccia As MXKit.cInterfaccia
Private WithEvents CFiltroDati As MXKit.CFiltro
Attribute CFiltroDati.VB_VarHelpID = -1
'======================================================
'           definizione variabili
'======================================================
'variabili di stato della form
Dim intSchOnTop1 As Integer
Dim intSchOnTop2 As Integer
Dim MlngButtonMask As Long
'variabili per visione
Dim mStrNomeVis As String 'nome visione
Dim mStrOrdinamento As String 'datafield colonna ordinamento
Dim mStrCriterio As String 'criterio visione primo livello
Dim mVntVariabili As Variant 'variabili da passare alla traccia
Dim mBolInizializza As Boolean 'risultato dell'inizializzazione
Dim mBolFiltro As Boolean 'la visione ha filtro si/no
Dim mStrPrevFiltro As String 'filtro precedentemente impostato sulla visione
Dim mStrCurFiltro As String 'filtro correntemente impostato sulla visione
Dim mBolImpDefault As Boolean

Dim mStrScriptVis As String

'RIF.A#11022
Private mSngMinWidth As Single
Private mSngMinHeight As Single
Private mSngCurrentWidth As Single
Private mSngCurrentHeight As Single

Public RibbonBar As RibbonBar

Public MOffSetTipoVis As Integer    'La Lista viene caricata su CLivello del Kit
Public MOffSetOpzioni As Integer
Dim MOffSetRaggruppa As Integer
Dim MOffSetAzioni As Integer
Dim MOffSetExportCrystal As Integer
Dim MOffSetExportOlap As Integer

Public Sub AbilitaRibbon(ByVal bolAbilita As Boolean)
    Dim i As Long
    For i = 1 To Me.CommandBars.ActiveMenuBar.Controls.Count
        Me.CommandBars.ActiveMenuBar.Controls(i).Enabled = bolAbilita
    Next i
    If SchedaTrovaBox.Visible And Me.CommandBars.VisualTheme = xtpThemeVisualStudio2008 Then  'Anomalia nr. 12213
        SchedaTrovaBox.Top = SchedaTrovaBox.Top - 375
        ssVisione(cTraccia.LivelloCorrente).Height = ssVisione(cTraccia.LivelloCorrente).Height - 375
        Frame(0).Height = Scheda(SCH_VISIONE).Height - Frame(0).Top - 60
    End If
End Sub

Public Sub Frm_SetOriginalSize()
    Me.Height = mSngMinHeight
    Me.width = mSngMinWidth
End Sub

Public Property Get ID_TIPOVIS() As Long
    ID_TIPOVIS = ID_VIS_TIPOVIS
End Property

Public Property Get ID_RAGGRUPPA() As Long
    ID_RAGGRUPPA = ID_VIS_RAGGRUPPA
End Property

Public Property Get ID_OPZIONIVIS() As Long
    ID_OPZIONIVIS = ID_VIS_OPZIONIVIS
End Property

Public Property Get ID_SEZFILTRO() As Long
    ID_SEZFILTRO = ID_VIS_FILTRO
End Property

Public Property Get ID_SEZVISIONE() As Long
    ID_SEZVISIONE = ID_VIS_VISIONE
End Property

Public Property Get ID_SEZTOTALI() As Long
    ID_SEZTOTALI = ID_VIS_TOTALI
End Property

Private Sub CFiltroDati_ValidazionePers(ByVal strNomeValid As String, ByVal strNomeCmpValid As String, bolEseguiValStd As Boolean, vntNewValore As Variant)
    Call ValidPersFiltri(strNomeValid, strNomeCmpValid, bolEseguiValStd, vntNewValore)
End Sub

Private Sub ComEsciTrova_Click()
    SchedaTrovaBox.Visible = False
    Call AbilitaRibbon(True)
    Call Form_Resize
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ID_VIS_FILTRO
            Call Ling_GotFocus(SCH_FILTRO)
        Case ID_VIS_VISIONE
            Call Ling_GotFocus(SCH_VISIONE)
        Case ID_VIS_TOTALI
            Call Ling_GotFocus(SCH_TOTALI)
        Case ID_VIS_TIPOVIS To ID_VIS_TIPOVIS + 20
            TxtTrova.text = ""     'Anomalia nr. 6530
            txtAvanzato.text = ""
            Call ssSpreadClear(ssTrova)
            cTraccia.pLivelloCorrente.strSQLTrova = ""
            cInterfaccia.IDListaComandoEvolus = Control.ID
            Call cInterfaccia.ListaClick(cTraccia, cTraccia.LivelloCorrente)
        Case ID_VIS_OPZIONIVIS To ID_VIS_OPZIONIVIS + 20
            cInterfaccia.IDListaComandoEvolus = Control.ID
            Call cInterfaccia.FiltroButtonClicked(cTraccia, cTraccia.LivelloCorrente, 1, Control.ID - ID_VIS_OPZIONIVIS, True)
        Case ID_VIS_AZIONI + 1 To ID_VIS_AZIONI + 20
            'eseguo l'azione associata
            Call cTraccia.EseguiAzione(Val(Control.Category))
        Case ID_VIS_RAGGRUPPA + 1 To ID_VIS_RAGGRUPPA + 20
            Call GestRaggruppa(Control.ID)
        Case ID_VIS_RAGGRUPPACHIUDI
            Call GestRaggruppa(ID_VIS_RAGGRUPPA + 2)
        Case ID_VIS_EXPORTCRYSTAL + 1 To ID_VIS_EXPORTCRYSTAL + 20
            Call GestExportCrystal(Control.ID)
        Case ID_VIS_EXPORTOLAP + 1 To ID_VIS_EXPORTOLAP + 20
            Call GestExportOlap(Control.ID)
        Case ID_VIS_EXPORTQLIK
            Call cTraccia.Export2QlikView
        Case ID_VIS_STAMPARAGGR
            Call cTraccia.pLivello(cTraccia.LivelloCorrente).GroupByPreview
            Call Form_Activate
        Case ID_RIBBON_EXPAND
            RibbonBar.Minimized = Not RibbonBar.Minimized
            Me.CommandBars.RecalcLayout
        Case ID_RIBBON_MINIMIZE
            RibbonBar.Minimized = Not RibbonBar.Minimized
            Me.CommandBars.RecalcLayout
        Case ID_VIS_TROVASTD, ID_VIS_TROVAAVZ
            If RibbonBar.FindControl(, ID_VIS_TROVA).Visible Or RibbonBar.Minimized Then
                Call AttivaTrova(Control.ID = ID_VIS_TROVAAVZ)
            End If
        Case ID_VIS_RICTOTALI
            If cTraccia.pGestTotali Then
                Call cTraccia.VisioneCaricaTotali(cTraccia.TotaleCorrente)
            End If
        Case Else
            cInterfaccia.IDComandiEvolus = Control.ID
            Call cInterfaccia.ComandiButtonClick(cTraccia, Nothing)
    End Select

End Sub

Private Sub GestExportCrystal(ByVal ID As Long)
    If ID = ID_VIS_EXPORTCRYSTAL + 1 Then
        If Not cTraccia.pLivello(cTraccia.LivelloCorrente).IsGroupBy Then
            Call cTraccia.Export2Crystal
        End If
    Else
        Dim ControlPopUp As CommandBarPopup
        Dim ControlItem As CommandBarControl
        
        Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_EXPORTCRYSTAL)
        Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, ID)
        If Not (ControlItem Is Nothing) Then
            If Not cTraccia.pLivello(cTraccia.LivelloCorrente).IsGroupBy Then
                Call cTraccia.Crystal_OpenReport(ControlItem.Category, ControlItem.Caption)
            End If
        End If
    End If
End Sub


Private Sub CommandBars_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.CommandBars.VisualTheme = xtpThemeVisualStudio2008 Then   'Anomalia nr. 12213
        Top = -375
    End If
End Sub

Private Sub CommandBars_Resize()
    Call Form_Resize
End Sub



Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ID_VIS_FILTRO: Control.Checked = (intSchOnTop1 = SCH_FILTRO)
        Case ID_VIS_VISIONE: Control.Checked = (intSchOnTop1 = SCH_VISIONE)
        Case ID_VIS_TOTALI: Control.Checked = (intSchOnTop1 = SCH_TOTALI)
        Case ID_RIBBON_EXPAND:
            Control.Visible = RibbonBar.Minimized
        Case ID_RIBBON_MINIMIZE:
            Control.Visible = Not RibbonBar.Minimized
    End Select
'    If intSchOnTop1 = SCH_VISIONE Then
'        'Altrimenti in caricamento della form di selezione il foglio non prende il focus
'        On Local Error Resume Next
'        If MbolLoading Then
'            If Me.Visible Then
'                ssVisione(cTraccia.LivelloCorrente).SetFocus
'                MbolLoading = False
'            End If
'        End If
'        On Local Error GoTo 0
'    End If
End Sub


Private Sub ctlImpFiltro_ImpostazioneDefaultCaricata()
    'Rif. Sviluppo Nr. 1041
    mBolImpDefault = True
End Sub

Private Sub cTraccia_EseguiAzionePersonale(ByVal intAzione As Integer, ByVal vntParametri As Variant, bolEseguiAzioneStd As Boolean)
#If TOOLS = 0 And ISNUCLEO = 0 Then
    Dim strNomeTraccia As String
    Dim intOrigineEvento As Integer
    Dim lngIDTesta As Long
    Dim lngIDRiga As Long
    Dim lngNumeroMov As Long
    Dim colParametri As Collection
    
    Set colParametri = New Collection
    strNomeTraccia = UCase(cTraccia.pNomeTraccia)

    Select Case strNomeTraccia
        Case "VIS_RIEPILOGOCCSTRUTT2", "VIS_RIEPILOGOCCSTRUTT_RIEP", "VIS_RIEPILOGOCCSTRUTT_DEST", "VIS_RIEPILOGOCCSTRUTTDEST_RIEP"
            intOrigineEvento = vntParametri(2)
            lngIDTesta = vntParametri(3)
            lngIDRiga = vntParametri(4)
            lngNumeroMov = 0
            
            If intAzione = 1 Then
                'necessario per non far perdere il focus alla linguetta corrente!
                Ling(0).Enabled = False: Ling(2).Enabled = False
                Call frmAssegnazioniCommCli.CaricaStruttura(vntParametri(1), vntParametri(2), vntParametri(3), vntParametri(4), True, cTraccia)
                Ling(0).Enabled = True: Ling(2).Enabled = True
            ElseIf intAzione = 2 Then
                GoSub ApriRiferimento
            End If
        
        Case "VIS_COMMCLI", "VIS_COMMCLI_RIEP", "VIS_RIEPILOGOCONSCOMM", "VIS_RIEPILOGOCONSCOMM_RIEP"
            If (strNomeTraccia = "VIS_COMMCLI") Or (strNomeTraccia = "VIS_COMMCLI_RIEP") Then
                On Error Resume Next
                lngNumeroMov = vntParametri(5)
                If Err.Number <> 0 Then lngNumeroMov = 0
                intOrigineEvento = vntParametri(2) 'Origine Evento
                On Error GoTo 0
            Else
                lngNumeroMov = 0
            End If
            
            On Error Resume Next
            lngIDTesta = vntParametri(3)
            lngIDRiga = vntParametri(4)
            If Err.Number <> 0 Then
                Call MXNU.MsgBoxEX(2800, vbInformation, Me.Caption) 'Evento di destinazione non assegnato o non disponibile per la visualizzazione corrente
                Exit Sub
            End If
            On Error GoTo 0
                    
            GoSub ApriRiferimento
    End Select
    Exit Sub

ApriRiferimento:
    Select Case intOrigineEvento
        Case Tipo_Documento
            Call FormLoader(frmGestioneDoc, 4000)
            DoEvents
            'assegno il valore dell'IDTesta all'indice uno della collection perchè
            'è lì che si aspetta di trovarlo la form dei documenti!
            Call colParametri.Add(lngIDTesta)
            Call MXNU.FrmMetodo.FormAttiva.AzioniMetodo(MetFDettVisione, colParametri)
        Case Tipo_PrimaNota
            Call FormLoader(frmPnContabile, 2400)
            DoEvents
            'assegno il valore dell'IDTesta all'indice uno della collection perchè
            'è lì che si aspetta di trovarlo la form prima nota contabile!
            Call colParametri.Add(lngIDTesta)
            Call MXNU.FrmMetodo.FormAttiva.AzioniMetodo(MetFDettVisione, colParametri)
        Case Tipo_OrdineProd
            If lngNumeroMov > 0 Then
                'rif. anomalia #7006
                Call MXNU.FrmMetodo.EseguiAzione("GestioneProdItem", 4, 5412)
                Dim colValoriChiave As Collection
                Set colValoriChiave = New Collection
                Call colValoriChiave.Add(lngIDTesta, "PROGRESSIVO")
                DoEvents 'necessario per beccare come form attiva la form dei documenti!
                Call MXNU.FrmMetodo.FormAttiva.AzioniMetodo(MetFDettVisione, colValoriChiave)
            Else
                Call FormLoader(frmOrdProd, 5410)
                DoEvents
                'assegno il valore dell'IDTesta all'indice uno della collection perchè
                'è lì che si aspetta di trovarlo la form ordini produzione!
                Call colParametri.Add(lngIDTesta)
                Call colParametri.Add(lngIDRiga)
                Call MXNU.FrmMetodo.FormAttiva.AzioniMetodo(MetFDettVisione, colParametri)
            End If
        Case Tipo_Calcolato
            Call MXNU.MsgBoxEX(2190, vbInformation, Me.Caption)
    End Select
    Return

#End If
End Sub

Private Sub cTraccia_PrimaEsecuzioneQuery(strQuery As String, bolSuccesso As Boolean)
    If StrComp(cTraccia.pNomeTraccia, "VIS_ANALISI_DISP", vbTextCompare) = 0 Then
        Dim vntDataElab As Variant
        'chiamata alla routine che esegue l'analisi e riempie la tabella temporanea
        'nota:  la strQuery della visione viene reimpostata assegnando alla parte "WHERE" solo l'IDSessione corrente!
        'Rif.sch. #7630 (gestione del filtro veloce- controllo se il filtro è stato modificato e solo in questo caso ri-eseguo)
        With cTraccia
            vntDataElab = .CFiltroDati.ParAgg("DataElab").ValoreFormula
            'Rif. sch. #7630 (passo la WHERE del Filtro (strCurFiltro) non la where di visione)
            Call CreaTempPerAnalisiDisp(strQuery, _
                                        mStrCurFiltro, _
                                        .pLivelloCorrente.SQLDammiORDERBY(.pLivelloCorrente.mIntVisione), _
                                        vntDataElab, _
                                        FiltroModificato(mStrCurFiltro))
        End With
    End If
End Sub

Private Sub CTraccia_RichiediRecordSetTotali(HrsTot As MXKit.CRecordSet, bolSuccesso As Boolean)
    
    Select Case UCase$(mStrNomeVis)
        Case "VIS_MOVMAG", "VIS_MOVMAG_BASE"
            'Anomalia nr. 4908
            On Local Error Resume Next
            If ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(5)) = OPTRA Then
                cTraccia.pTotale(cmbTotali.listIndex + 1).colGruppi(cTraccia.pTotale(cmbTotali.listIndex + 1).colGruppi.Count).colTotIniziali(2).strOmettiSe = "(%PARZIALE1%=0 AND %PARZIALE2%=0 AND %PARZIALE3%=0 AND %PARZIALE4%=0 AND %PARZIALE6%=0 AND %PARZIALE7%=0)"
            Else
                cTraccia.pTotale(cmbTotali.listIndex + 1).colGruppi(cTraccia.pTotale(cmbTotali.listIndex + 1).colGruppi.Count).colTotIniziali(2).strOmettiSe = "-1"
            End If
            On Local Error GoTo 0
        
            Call Totali_AddRecordIniziali(cTraccia, HrsTot, False, cmbTotali.listIndex + 1, 4, 5, ssFiltroDati, False)
        Case "VIS_MOVCON"
            'Anomalia nr. 4908
            On Local Error Resume Next
            If ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(3)) = OPTRA Then
                cTraccia.pTotale(cmbTotali.listIndex + 1).colGruppi(2).colTotIniziali(2).strOmettiSe = "(%PARZIALE3%=0)"
            Else
                cTraccia.pTotale(cmbTotali.listIndex + 1).colGruppi(2).colTotIniziali(2).strOmettiSe = "-1"
            End If
            On Local Error GoTo 0
    
            Call Totali_AddRecordIniziali(cTraccia, HrsTot, False, cmbTotali.listIndex + 1, 4, 3, ssFiltroDati, False, , 5, 2)
            
            'Anomalia nr. 3233
            Dim cGrp As MXKit.CGruppo
            Dim cPrz As MXKit.CTotAgg
            Dim strQuery As String
            
            For Each cGrp In cTraccia.pTotale(cmbTotali.listIndex + 1).colGruppi
                'If InStr(cGrp.strNomeGruppo, "Esercizio") > 0 Then
                If InStr(LCase(cGrp.CColGruppo.strDataField), "esercizio") > 0 Then   'Rif. Anomalia 7273: sostituito strNomegruppo con CColGruppo.strDataField
                    Set cPrz = cGrp.colTotIniziali(1)
                    strQuery = cPrz.colQueryTotali(3)
                    strQuery = Replace(strQuery, "%GRUPPO2%-1", "%GRUPPO2%")
                    strQuery = Replace(strQuery, "VistaMovCont", "%VISTA%")
                    strQuery = Replace(strQuery, "VISTASALDIFINALIPN", "%VISTA%")
                    strQuery = Replace(strQuery, "VISTASALDIINIZIALIPN", "%VISTA%")
                    'Anomalia nr. 4898
                    strQuery = Replace(strQuery, "VISTACONTIPATRIMONIALI", "%VISTA%")
                    Select Case MenmTipoSaldi
                        Case enmMovApertura
                            strQuery = Replace(strQuery, "%VISTA%", "VISTASALDIINIZIALIPN")
                            cPrz.strTitolo = MXNU.CaricaStringaRes(205091)
                        Case enmMovChiusuraAP
                            strQuery = Replace(strQuery, "%VISTA%", "VISTASALDIFINALIPN")
                            strQuery = Replace(strQuery, "%GRUPPO2%", "%GRUPPO2%-1")
                            cPrz.strTitolo = MXNU.CaricaStringaRes(204993)
                        Case enmSommaSaldiAP
                            strQuery = Replace(strQuery, "%VISTA%", "VISTACONTIPATRIMONIALI")
                            strQuery = Replace(strQuery, "%GRUPPO2%", "%GRUPPO2%-1")
                            cPrz.strTitolo = MXNU.CaricaStringaRes(204992)
                    End Select
                    Call cPrz.colQueryTotali.Remove(3)
                    Call cPrz.colQueryTotali.Add(strQuery)
                End If
            Next cGrp
    End Select
    bolSuccesso = True

End Sub
Private Sub ControllaMovApCh()
    Dim strSQL As String
    Dim q As Integer
    Dim hndtn As MXKit.CRecordSet
    Dim intEsercizio As Integer
    
    intEsercizio = ssCellGetValue(ssFiltroDati, COLVALOREDA, CFiltroDati.IdFiltro2Riga(4))
    
    'Esistenza Mov. di Apertura / Chiusura
    MenmTipoSaldi = enmSommaSaldiAP
    
    strSQL = "select TOP 1 Progressivo FROM TesteContabilita WHERE Causale=" & MXNU.Vincoli(CAUS_APERTURA) & " AND Esercizio=" & intEsercizio
    Set hndtn = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    If Not MXDB.dbFineTab(hndtn) Then
        MenmTipoSaldi = enmMovApertura
    End If
    q = MXDB.dbChiudiSS(hndtn)
    
    If MenmTipoSaldi = enmSommaSaldiAP Then
        strSQL = "select TOP 1 Progressivo FROM TesteContabilita WHERE Causale=" & MXNU.Vincoli(CAUS_CHIUSURA) & " AND Esercizio=" & intEsercizio - 1
        Set hndtn = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        If Not MXDB.dbFineTab(hndtn) Then
            MenmTipoSaldi = enmMovChiusuraAP
        End If
        q = MXDB.dbChiudiSS(hndtn)
    End If

End Sub

'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
'           EVENTI DELLA FORM
'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
Private Sub Form_Activate()
    Call MXNU.Attiva_Toolbar(hwnd, MlngButtonMask)
    Call MXNU.ImpostaFormAttiva(Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And SchedaTrovaBox.Visible Then
        SchedaTrovaBox.Visible = False
        Call AbilitaRibbon(True)
        Call Form_Resize
    End If
End Sub


Private Sub Form_Load()
    Me.HelpContextID = FormProp.FormID
    MlngButtonMask = BTN_STP_MASK
    Dim strDim As String
    
    'Leggo le dimensioni originali della form dal file MWForm.Ini
    strDim = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MWForm.ini", Me.NAME, "Height", "")
    If strDim <> "" Then
        Me.Height = Val(strDim)
    End If
    strDim = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MWForm.ini", Me.NAME, "Width", "")
    If strDim <> "" Then
        Me.width = Val(strDim)
    End If
    '...e le memorizzo
    mSngMinHeight = Me.Height
    mSngMinWidth = Me.width
    mSngCurrentHeight = mSngMinHeight
    mSngCurrentWidth = mSngMinWidth
       
End Sub

' rif.sch. A5149
' controllato che la visione non sia in caricamento
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Not (cTraccia Is Nothing) Then
        On Local Error Resume Next
        If cTraccia.InElaborazione Or cTraccia.InEsecuzioneAzione Then   'Anomalia nr. 11213
            Cancel = True
        End If
        If ssVisione(cTraccia.LivelloCorrente).SheetCount > 1 Then
            Call cTraccia.pLivello(cTraccia.LivelloCorrente).GroupByClose
        End If
        On Local Error GoTo 0
    End If
'Per Metodo Evolus
'*****************************************************************
'If Not Cancel Then
'        On Local Error Resume Next
'        If (Not mResize Is Nothing) Then
'                mResize.Terminate
'                Set mResize = Nothing
'        End If
'        On Local Error GoTo 0
'End If
'*****************************************************************
End Sub

Private Sub Form_Resize()
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim sngDeltaWidth As Single
    Dim sngDeltaHeight As Single
    Dim i As Integer

    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long

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

    
    Me.CommandBars.GetClientRect Left, Top, Right, Bottom
    
    If Right >= Left And Bottom >= Top Then
        'RIF.A#11022 - gestione manuale del resize
            Scheda(SCH_FILTRO).Move 0, Top, sngWidth, Bottom - Top
            Scheda(SCH_VISIONE).Move 0, Top, sngWidth, Bottom - Top
            Scheda(SCH_TOTALI).Move 0, Top, sngWidth, Bottom - Top
            For i = SSCH2_VALORI To SSCH2_GRAFICO
                Scheda(i).Height = Scheda(SCH_TOTALI).Height - Scheda(i).Top - 60
                Scheda(i).width = Scheda(i).width + sngDeltaWidth
            Next i
        '***SchedaTrovaBox.Top = SchedaTrovaBox.Top + sngDeltaHeight '- 50
        '***SchedaTrovaBox.width = SchedaTrovaBox.width + sngDeltaWidth
        'Frame(0).Height = Frame(0).Height + sngDeltaHeight
        Frame(0).Height = Scheda(SCH_VISIONE).Height - Frame(0).Top - 60
        Frame(0).width = Frame(0).width + sngDeltaWidth
        Dim idx As Integer
        For idx = 0 To cTraccia.pNumeroLivelli
            'ssVisione(idx).Height = ssVisione(idx).Height + sngDeltaHeight
            ssVisione(idx).Height = Frame(0).Height - ssVisione(idx).Top - 60
            ssVisione(idx).width = ssVisione(idx).width + sngDeltaWidth
            If MXNU.ResizeProporzionale Then
                If (Me.Height - mSngMinHeight) > (mSngMinHeight * MXNU.PercResizeProporzionale / 100) Then ssSpreadSetFontSize ssVisione(idx), 10 Else ssSpreadSetFontSize ssVisione(idx), 8
            End If
            If SchedaTrovaBox.Visible Then
                ssVisione(idx).Height = ssVisione(idx).Height - SchedaTrovaBox.Height - 90
            End If
        Next
        
        ssFiltroDati.Height = Scheda(SCH_FILTRO).Height - ssFiltroDati.Top - 60
        'ssFiltroDati.Height = ssFiltroDati.Height + sngDeltaHeight
        ssFiltroDati.width = ssFiltroDati.width + sngDeltaWidth
        'ssTotali.Height = ssTotali.Height + sngDeltaHeight
        ssTotali.Height = Scheda(SSCH2_VALORI).Height - ssTotali.Top - 60
        ssTotali.width = ssTotali.width + sngDeltaWidth
    
        'objChart.Height = objChart.Height + sngDeltaHeight
        objChart.Height = Scheda(SSCH2_GRAFICO).Height - objChart.Top - 60
        objChart.width = objChart.width + sngDeltaWidth
        
        
        If SchedaTrovaBox.Visible Then
            SchedaTrovaBox.Top = Me.Height - SchedaTrovaBox.Height - Top - 300
        End If
        
        If Not (cTraccia Is Nothing) Then
            Call cTraccia.CalcVisibleRows
        End If
        
        If Me.CommandBars.VisualTheme = xtpThemeVisualStudio2008 Then   'Anomalia nr. 12213
            SchedaTrovaBox.Top = SchedaTrovaBox.Top - 375
        End If
    End If
    
'    For i = SCH_FILTRO To SSCH2_GRAFICO
'        Scheda(i).Height = Scheda(i).Height + sngDeltaHeight
'        Scheda(i).width = Scheda(i).width + sngDeltaWidth
'    Next i
'
'    ssFiltroDati.Height = ssFiltroDati.Height + sngDeltaHeight
'    ssFiltroDati.width = ssFiltroDati.width + sngDeltaWidth
'
'    For i = 0 To cTraccia.pNumeroLivelli
'        ssVisione(i).Height = ssVisione(i).Height + sngDeltaHeight
'        ssVisione(i).width = ssVisione(i).width + sngDeltaWidth
'        cmbVisione(i).Left = cmbVisione(i).Left + sngDeltaWidth
'        If MXNU.ResizeProporzionale Then
'            If (Me.Height - mSngMinHeight) > (mSngMinHeight * MXNU.PercResizeProporzionale / 100) Then ssSpreadSetFontSize ssVisione(i), 10 Else ssSpreadSetFontSize ssVisione(i), 8
'        End If
'    Next
'
'    Frame(0).Height = Frame(0).Height + sngDeltaHeight
'    Frame(0).width = Frame(0).width + sngDeltaWidth
'    SchedaTrovaBox.Top = SchedaTrovaBox.Top + sngDeltaHeight
'    SchedaTrovaBox.width = SchedaTrovaBox.width + sngDeltaWidth
'
'    ssTotali.Height = ssTotali.Height + sngDeltaHeight
'    ssTotali.width = ssTotali.width + sngDeltaWidth
'
'    objChart.Height = objChart.Height + sngDeltaHeight
'    objChart.width = objChart.width + sngDeltaWidth

End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = vbNormal Then
        Call MXNU.ScriviProfilo(MXNU.PercorsoPreferenze & "\MWForm" & MXNU.NTerminale & ".ini", "FRMVISIONI_" & Me.HelpContextID, "Width", Me.width)
        Call MXNU.ScriviProfilo(MXNU.PercorsoPreferenze & "\MWForm" & MXNU.NTerminale & ".ini", "FRMVISIONI_" & Me.HelpContextID, "Height", Me.Height)
    End If
    If Not cTraccia Is Nothing Then
        Call cTraccia.TerminaVisione
        Set cTraccia = Nothing
        Set cInterfaccia = Nothing
    End If
    ctlImpFiltro.Termina
    Set CFiltroDati = Nothing
    Set FormProp = Nothing
    Set frmVisioni = Nothing
End Sub


Private Sub Ling_GotFocus(Index As Integer)
Dim intOldLing As Integer
Static bolCaricataVisione As Boolean
    If (Index <> intSchOnTop1 And Index <> intSchOnTop2) Then
        Dim ControlPopUp As CommandBarPopup
        Dim ControlItem As CommandBarControl
        Select Case Index
            Case SCH_FILTRO
                intOldLing = intSchOnTop1
                intSchOnTop1 = Index
                RibbonBar.FindGroup(ID_VISGROUP_GROUPVIS).Visible = True
                RibbonBar.FindControl(, ID_VIS_TIPOVIS).Visible = False
                RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = False
                RibbonBar.FindControl(, ID_VIS_RICARICA).Visible = False
                RibbonBar.FindControl(, ID_VIS_STAMPARAGGR).Visible = False
                RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = False
                RibbonBar.FindControl(, ID_VIS_TROVA).Visible = False
                Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA)
                Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, ID_VIS_RAGGRUPPA + 1)
                If Not (ControlItem Is Nothing) Then ControlItem.Enabled = False
                Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, ID_VIS_RAGGRUPPA + 2)
                If Not (ControlItem Is Nothing) Then ControlItem.Enabled = False
                RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = False
                RibbonBar.FindGroup(ID_VISGROUP_GROUPTOTALI).Visible = False
                RibbonBar.FindControl(, ID_VIS_RAGGRUPPACHIUDI).Visible = False
            Case SCH_VISIONE
                intOldLing = intSchOnTop1
                intSchOnTop1 = Index
                Ling(SCH_FILTRO).Enabled = False
                GoSub caricaVisione
                Ling(SCH_FILTRO).Enabled = True
                'cmbVisione(cTraccia.LivelloCorrente).ZOrder 0
                RibbonBar.FindGroup(ID_VISGROUP_GROUPVIS).Visible = True
                cmbVisione(cTraccia.LivelloCorrente).Visible = False
                RibbonBar.FindControl(, ID_VIS_TIPOVIS).Visible = True
                If cTraccia.LivelloCorrente > 0 Then
                    RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = (cTraccia.colLivelli(cTraccia.LivelloCorrente).CFiltro.colOpzioni.Count > 0)
                End If
                RibbonBar.FindControl(, ID_VIS_RICARICA).Visible = True
                RibbonBar.FindControl(, ID_VIS_RAGGRUPPA).Visible = True
                If cTraccia.LivelloCorrente > 0 Then
                    If cTraccia.pLivello(cTraccia.LivelloCorrente).ssVisione.SheetCount > 1 Then
                        Select Case cTraccia.pLivello(cTraccia.LivelloCorrente).ssVisione.ActiveSheet
                            Case 1
                                RibbonBar.FindGroup(ID_VISGROUP_GROUPVIS).Visible = True
                                RibbonBar.FindControl(, ID_VIS_TIPOVIS).Visible = False
                                RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = (cTraccia.colLivelli(cTraccia.LivelloCorrente).CFiltro.colOpzioni.Count > 0)
                                RibbonBar.FindControl(, ID_VIS_RICARICA).Visible = False
                                RibbonBar.FindControl(, ID_VIS_STAMPARAGGR).Visible = True
                                RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = True
                                RibbonBar.FindControl(, ID_VIS_EXPORTCRYSTAL).Visible = False
                                RibbonBar.FindControl(, ID_VIS_EXPORTOLAP).Visible = False
                                RibbonBar.FindControl(, ID_VIS_EXPORTOOCALC).Visible = False
                                RibbonBar.FindControl(, ID_VIS_EXPORTQLIK).Visible = False
                                RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = False
                                RibbonBar.FindControl(, ID_VIS_TROVA).Visible = False
                            Case 2
                                RibbonBar.FindGroup(ID_VISGROUP_GROUPVIS).Visible = True
                                RibbonBar.FindControl(, ID_VIS_TIPOVIS).Visible = True
                                RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = (cTraccia.colLivelli(cTraccia.LivelloCorrente).CFiltro.colOpzioni.Count > 0)
                                RibbonBar.FindControl(, ID_VIS_RICARICA).Visible = True
                                RibbonBar.FindControl(, ID_VIS_STAMPARAGGR).Visible = False
                                RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = True
                                RibbonBar.FindControl(, ID_VIS_EXPORTCRYSTAL).Visible = True
                                RibbonBar.FindControl(, ID_VIS_EXPORTOLAP).Visible = True
                                RibbonBar.FindControl(, ID_VIS_EXPORTOOCALC).Visible = True
                                RibbonBar.FindControl(, ID_VIS_EXPORTQLIK).Visible = True
                                RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = True
                                RibbonBar.FindControl(, ID_VIS_TROVA).Visible = True
                            Case 3
                                RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = False
                                RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = False
                                RibbonBar.FindGroup(ID_VISGROUP_GROUPVIS).Visible = False
                                RibbonBar.FindControl(, ID_VIS_TROVA).Visible = False
                        End Select
                        RibbonBar.FindControl(, ID_VIS_RAGGRUPPACHIUDI).Visible = True
                    Else
                        RibbonBar.FindGroup(ID_VISGROUP_GROUPVIS).Visible = True
                        RibbonBar.FindControl(, ID_VIS_TIPOVIS).Visible = True
                        RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = (cTraccia.colLivelli(cTraccia.LivelloCorrente).CFiltro.colOpzioni.Count > 0)
                        RibbonBar.FindControl(, ID_VIS_RICARICA).Visible = True
                        RibbonBar.FindControl(, ID_VIS_STAMPARAGGR).Visible = False
                        RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = True
                        RibbonBar.FindControl(, ID_VIS_EXPORTCRYSTAL).Visible = True
                        RibbonBar.FindControl(, ID_VIS_EXPORTOLAP).Visible = True
                        RibbonBar.FindControl(, ID_VIS_EXPORTOOCALC).Visible = True
                        RibbonBar.FindControl(, ID_VIS_EXPORTQLIK).Visible = True
                        RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = True
                        RibbonBar.FindControl(, ID_VIS_TROVA).Visible = True
                        RibbonBar.FindControl(, ID_VIS_RAGGRUPPACHIUDI).Visible = False
                    End If
                End If
                RibbonBar.FindGroup(ID_VISGROUP_GROUPTOTALI).Visible = False
                'RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = True
                'RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = True
                Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA)
                Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, ID_VIS_RAGGRUPPA + 1)
                If Not (ControlItem Is Nothing) Then ControlItem.Enabled = True
                Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, ID_VIS_RAGGRUPPA + 2)
                If Not (ControlItem Is Nothing) Then ControlItem.Enabled = True
                Call LoadAzioni
            Case SCH_TOTALI
                intOldLing = intSchOnTop1
                intSchOnTop1 = Index
                intSchOnTop2 = SSCH2_GRAFICO
                
                'Anomalia nr. 3233
                If StrComp(mStrNomeVis, "VIS_MOVCON", vbTextCompare) = 0 Then
                    Call ControllaMovApCh
                End If

                Call Ling_GotFocus(SSCH2_VALORI)
                'RibbonBar.FindControl(, ID_VIS_TIPOVIS).Visible = False
                'RibbonBar.FindControl(, ID_VIS_RAGGRUPPA).Visible = False
                RibbonBar.FindGroup(ID_VISGROUP_GROUPVIS).Visible = False
                RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = False
                RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = False
                RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = True
                RibbonBar.FindControl(, ID_VIS_EXPORTCRYSTAL).Visible = False
                RibbonBar.FindControl(, ID_VIS_EXPORTOLAP).Visible = False
                RibbonBar.FindControl(, ID_VIS_EXPORTEXCEL).Visible = True
                RibbonBar.FindControl(, ID_VIS_EXPORTOOCALC).Visible = True
                RibbonBar.FindControl(, ID_VIS_EXPORTQLIK).Visible = True
                RibbonBar.FindControl(, ID_VIS_TROVA).Visible = False
                RibbonBar.FindGroup(ID_VISGROUP_GROUPTOTALI).Visible = True
                RibbonBar.FindControl(, ID_VIS_RAGGRUPPACHIUDI).Visible = False
            Case SSCH2_VALORI
                'carico i totali se selezionati e se non vengo dal grafico
                If Not Ling(SSCH2_GRAFICO).OnTop Then
                    GoSub caricaVisione
                    With cTraccia
                        If (.TotaleCorrente > 0) Then
                            If .pRicalcolaTotali Or FiltroModificato(mStrCurFiltro) Then
                                'cancello il foglio dei totali per evidenziare che devono essere ricalcolati (rif.sch.1572)
                                ssSpreadClear ssTotali, SS_CLR_TEXTONLY
'                                If MXNU.MsgBoxEX(1562, vbQuestion + vbYesNo + vbDefaultButton2, 1007) = vbYes Then
'                                    Call .VisioneCaricaTotali(.TotaleCorrente)
'                                End If
                                'Anomalia nr. 11587
                                Call ssSpreadImposta(ssTotali)
                                Call ssCellSetBorder(ssTotali, -1, -1, SS_BORDER_STYLE_BLANK, SS_BORDER_TYPE_NONE, vbButtonFace, vbButtonText)
                            End If
                        End If
                    End With
                End If
                intOldLing = intSchOnTop2
                intSchOnTop2 = Index
            Case SSCH2_GRAFICO
                intOldLing = intSchOnTop2
                intSchOnTop2 = Index
                If Not cTraccia.pGestGrafico Then
                    GoTo fine_LingGotFocus
                Else
                    If (Not cTraccia.pChart.DisegnaGrafico) Then
                        GoTo fine_LingGotFocus
                    End If
                End If
            Case Else
                intOldLing = intSchOnTop1
                intSchOnTop1 = Index
        End Select
        
        DoEvents
        Scheda(Index).Visible = True
        Scheda(intOldLing).Visible = False
        Ling(intOldLing).OnTop = False
        Scheda(Index).ZOrder vbBringToFront
        Ling(Index).OnTop = True
    End If
    
    'Per Metodo Evolus
Call CambiaZOrderLinguette(Me)
    Me.KeyPreview = (intSchOnTop1 <> SCH_FILTRO)
    cTraccia.TrovaKeyPreview = (intSchOnTop1 = SCH_VISIONE)
    DoEvents
    
        
fine_LingGotFocus:
    Exit Sub

caricaVisione:
    'rif.sch. A5149
    If Not (cTraccia Is Nothing) Then
        With cTraccia
            If (Not .colLivelli(1).bolImpostato) Then
                'prima volta->imposto livello 1
                .LivelloCorrente = 1
                If .pGestFiltroDati Then mStrCurFiltro = CFiltroDati.SQLFiltro
                Call .CalcVisibleRows
            ElseIf FiltroModificato(mStrCurFiltro) Then
                'ricarico la visione
                If .pLivello(.LivelloCorrente).IsGroupBy Then
                    Call .pLivello(.LivelloCorrente).GroupBy(.pLivello(.LivelloCorrente).GroupByName)
                Else
                    Call .VisioneCaricaDati(1, .colLivelli(1).mIntVisione)
                End If
            End If
            mStrPrevFiltro = mStrCurFiltro
        End With
    End If
Return

End Sub

Private Sub cmbVisione_Click(Index As Integer)
    Call cInterfaccia.ListaClick(cTraccia, Index)
End Sub

Private Sub ssVisione_BeforeUserSort(Index As Integer, ByVal Col As Long, ByVal State As FPSpreadADO.BeforeUserSortStateConstants, DefaultAction As FPSpreadADO.BeforeUserSortDefaultActionConstants)
    DefaultAction = BeforeUserSortDefaultActionCancel
End Sub

Private Sub ssVisione_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Call cInterfaccia.VisioneButtonClicked(cTraccia, Index, Col, Row, ButtonDown)
    If (ssVisione(Index).ActiveSheet = 3) And (Row = 1) And (Col = 3 Or Col = 4) Then Call LoadGroupBy(True)
End Sub

Private Sub ssVisione_Change(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Call cInterfaccia.VisioneChange(cTraccia, Index, Col, Row)
End Sub

Private Sub ssVisione_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Call cInterfaccia.VisioneClick(cTraccia, Index, Col, Row)
End Sub

Private Sub ssVisione_ColWidthChange(Index As Integer, ByVal Col1 As Long, ByVal Col2 As Long)
    Call cInterfaccia.VisioneColWidthChange(cTraccia, Index, Col1)
End Sub

Private Sub ssVisione_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Call cInterfaccia.VisioneDblClick(cTraccia, Index, Col, Row)
End Sub

Private Sub ssVisione_KeyPress(Index As Integer, KeyAscii As Integer)
    Call cInterfaccia.VisioneKeyPress(cTraccia, Index, KeyAscii)
End Sub

Private Sub ssVisione_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call cInterfaccia.VisioneLeaveCell(cTraccia, Index, Col, Row, NewCol, NewRow, Cancel)
End Sub

Private Sub ssVisione_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call cInterfaccia.VisioneMouseDown(cTraccia, Index, Button, Shift, x, y)
End Sub

Private Sub ssVisione_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call cInterfaccia.VisioneMouseMove(cTraccia, Index, Button, Shift, x, y)
End Sub

Private Sub ssOpzioni_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If cTraccia.MBolVisioneRibbon Then
        cInterfaccia.IDListaComandoEvolus = ID_VIS_OPZIONIVIS + Row
    End If
    Call cInterfaccia.FiltroButtonClicked(cTraccia, Index, Col, Row, ButtonDown)
End Sub

Private Sub ssVisione_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Call cInterfaccia.VisioneDragDrop(cTraccia, Index, Source, x, y)
End Sub

Private Sub ssVisione_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Call cInterfaccia.VisioneDragOver(cTraccia, Index, Source, x, y, State)
End Sub

Private Sub ssVisione_SheetChanged(Index As Integer, ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    Call cInterfaccia.VisioneSheetChanged(cTraccia, Index, OldSheet, NewSheet)
End Sub

Private Sub ssVisione_SheetChanging(Index As Integer, ByVal OldSheet As Integer, ByVal NewSheet As Integer, Cancel As Variant)
    Call cInterfaccia.VisioneSheetChanging(cTraccia, Index, OldSheet, NewSheet, Cancel)
End Sub

Private Sub ssVisione_TopLeftChange(Index As Integer, ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    Call cInterfaccia.VisioneTopLeftChange(cTraccia, Index, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Private Sub Scheda_Paint(Index As Integer)
    Call SchedaOmbreggiaControlli(Scheda(Index))
End Sub

'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
'           PROPRIETA' PUBBLICHE DELLA FORM
'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
Public Property Let pCriterio(new_Valore As String)
    mStrCriterio = new_Valore
End Property

Public Property Get pCriterio() As String
    pCriterio = mStrCriterio
End Property

Public Property Let pNomeVisione(new_Valore As String)
    mStrNomeVis = new_Valore
End Property

Public Property Get pNomeVisione() As String
    pNomeVisione = mStrNomeVis
End Property

Public Property Let pOrdinamento(new_Valore As String)
    mStrOrdinamento = new_Valore
End Property

Public Property Get pOrdinamento() As String
    pOrdinamento = mStrOrdinamento
End Property

Public Property Let pScriptVisione(new_Valore As String)
    mStrScriptVis = new_Valore
End Property

Public Property Get pScriptVisione() As String
    pScriptVisione = mStrScriptVis
End Property


'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
'           FUNZIONI PRIVATE DELLA FORM
'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&

Private Function LoadVisione() As Boolean
Dim intCurLiv As Integer

    metodo.MousePointer = vbHourglass
    Call MXNU.LeggiRisorseControlli(Me)
    If mStrScriptVis <> "" Then
        Call cTraccia.pDefTraccia.AssegnaScriptDefinizione(mStrScriptVis)
    End If
    mBolInizializza = cTraccia.Inizializza(mStrNomeVis, _
        MXKit.tivVisione, _
        mStrOrdinamento, _
        mStrCriterio, _
        MXKit.selSelezioneStandard, _
        Nothing)
    If (mBolInizializza) Then
        cTraccia.MBolVisioneRibbon = True
        Call cTraccia.ImpostaVariabili(mVntVariabili)
        'disabilito la gestione della visione paginata per evitare che venga reimpotata la where del filtro (rif.sch.2130)
        If StrComp(mStrNomeVis, "VIS_ANALISI_DISP", vbTextCompare) = 0 Then
            cTraccia.pVisionePaginata = False
        End If
        'imposto il foglio del filtro
        Set CFiltroDati = MXFT.CreaCFiltro()
        mBolFiltro = (CFiltroDati.InizializzaFiltro(cTraccia.colLivelli(1).strFiltro, ssFiltroDati))
        If (mBolFiltro) Then
            mBolImpDefault = False
            Call ctlImpFiltro.Inizializza(MXDB, MXNU, MXVI, cTraccia.colLivelli(1).strFiltro, CFiltroDati, Nothing, ssFiltroDati, hndDBArchivi, GIMP_FILTRONORMALE)
            Set cTraccia.CFiltroDati = CFiltroDati
        End If
        'imposto i controlli della traccia
        Call cTraccia.ImpostaControlli(Me, , _
            TxtTrova, txtAvanzato, ssTrova, comTrova, comAvanzato, lstHelp, _
            ssTotali, cmbTotali, _
            objChart, _
            ssFiltroDati)
        'carico i controlli
        For intCurLiv = 1 To cTraccia.colLivelli.Count
            Load ssVisione(intCurLiv)
            Load cmbVisione(intCurLiv)
            Load ssOpzioni(intCurLiv)
            Call cTraccia.colLivelli(intCurLiv).ImpostaControlli(ssVisione(intCurLiv), cmbVisione(intCurLiv), ssOpzioni(intCurLiv))
        Next
        'imposto le schede
        '***Call Inizializza_Schede(Me, 3)
        Call CreateRibbon
        
        
        intSchOnTop1 = SCH_TOTALI
        intSchOnTop2 = SSCH2_GRAFICO
        
        '******GESTIONE SCHEDE SU RIBBON *************************************************
        CommandBars.FindControl(, ID_VIS_FILTRO).Visible = mBolFiltro
        Scheda(SCH_FILTRO).Visible = mBolFiltro
        CommandBars.FindControl(, ID_VIS_TOTALI).Visible = cTraccia.pGestTotali
        Scheda(SCH_TOTALI).Visible = cTraccia.pGestTotali
        'Ling(SCH_FILTRO).Visible = mBolFiltro
        'Scheda(SCH_FILTRO).Visible = mBolFiltro
        'Ling(SCH_TOTALI).Visible = cTraccia.pGestTotali
        'Scheda(SCH_TOTALI).Visible = cTraccia.pGestTotali
        Ling(SSCH2_GRAFICO).Visible = cTraccia.pGestGrafico
        Scheda(SSCH2_GRAFICO).Visible = cTraccia.pGestGrafico
        '*********************************************************************************
        
        If mBolFiltro Then
            'Call Ling_GotFocus(SCH_FILTRO)
            Call CommandBars_Execute(CommandBars.FindControl(, ID_VIS_FILTRO))
        Else
            'Ling(SCH_VISIONE).Left = 0
            'Ling(SCH_TOTALI).Left = Ling(SCH_VISIONE).width
            'Call Ling_GotFocus(SCH_VISIONE)
            Call CommandBars_Execute(CommandBars.FindControl(, ID_VIS_VISIONE))
        End If
        
        Call LoadBrioModels
        Call LoadCrystalReports
        Call LoadGroupBy
        
        'mostro la finestra
        Call CentraFinestra(Me.hwnd)
        metodo.MousePointer = vbDefault
    Else
        metodo.MousePointer = vbDefault
        Call MXNU.MsgBoxEX(MXNU.CaricaStringaRes(1050, mStrNomeVis), vbCritical, App.EXEName)
    End If
     
    ''''GESTIONE ACCESSI VISIONI'''''' RIF. AN.#8772 RZ
    Dim intAccLingFiltro As Integer
    Dim intAccLingVisione As Integer
    Dim intAccLingTotali As Integer
    Dim intAccLingXXXX As Integer
    Dim intAccLingGrafico As Integer
    
    If MXNU.CtrlAccessi Then
        'filtro
        intAccLingFiltro = LeggiAccessi(MXNU.UtenteAttivo, Me.HelpContextID, 0, False)
        Call ImpostaAccessiVisione(intAccLingFiltro, 0)
        'visione
        intAccLingVisione = LeggiAccessi(MXNU.UtenteAttivo, Me.HelpContextID, 1, False)
        Call ImpostaAccessiVisione(intAccLingVisione, 1)
        'totali
        intAccLingTotali = LeggiAccessi(MXNU.UtenteAttivo, Me.HelpContextID, 2, False)
        Call ImpostaAccessiVisione(intAccLingTotali, 2)
        'situazione
        intAccLingXXXX = LeggiAccessi(MXNU.UtenteAttivo, Me.HelpContextID, 3, False)
        Call ImpostaAccessiVisione(intAccLingXXXX, 3)
        'grafico
        intAccLingGrafico = LeggiAccessi(MXNU.UtenteAttivo, Me.HelpContextID, 4, False)
        Call ImpostaAccessiVisione(intAccLingGrafico, 4)
    End If
    
    'risultato
    LoadVisione = mBolInizializza
End Function

''''GESTIONE ACCESSI VISIONI'''''' RIF. AN.#8872 RZ
''''ho tenuto separata la gestione lettura e modifica pur avendo in questo contesto lo stesso comportamento
''''nel caso in cui vengano effettuate altre richieste specifiche
Private Sub ImpostaAccessiVisione(intAcc As Integer, LingID As Integer)

    Select Case LingID
        Case 0
            If intAcc = ACC_NESSUNO Then
                If Not RibbonBar.FindControl(xtpControlButton, ID_VIS_FILTRO) Is Nothing Then RibbonBar.FindControl(xtpControlButton, ID_VIS_FILTRO).Visible = False
            ElseIf (intAcc > ACC_NESSUNO) And (intAcc <= ACC_TUTTI) Then
                If (intAcc And ACC_LETTURA) Then
                    If Not RibbonBar.FindControl(xtpControlButton, ID_VIS_FILTRO) Is Nothing Then RibbonBar.FindControl(xtpControlButton, ID_VIS_FILTRO).Enabled = True
                ElseIf (intAcc And ACC_MODIFICA) Then
                    If Not RibbonBar.FindControl(xtpControlButton, ID_VIS_FILTRO) Is Nothing Then RibbonBar.FindControl(xtpControlButton, ID_VIS_FILTRO).Visible = True
                End If
            End If
        Case 1
            If intAcc = ACC_NESSUNO Then
                If Not RibbonBar.FindControl(xtpControlButton, ID_VIS_VISIONE) Is Nothing Then RibbonBar.FindControl(xtpControlButton, ID_VIS_VISIONE).Visible = False
            ElseIf (intAcc > ACC_NESSUNO) And (intAcc <= ACC_TUTTI) Then
                If (intAcc And ACC_LETTURA) Then
                    If Not RibbonBar.FindControl(xtpControlButton, ID_VIS_VISIONE) Is Nothing Then RibbonBar.FindControl(xtpControlButton, ID_VIS_VISIONE).Enabled = True
                ElseIf (intAcc And ACC_MODIFICA) Then
                    If Not RibbonBar.FindControl(xtpControlButton, ID_VIS_VISIONE) Is Nothing Then RibbonBar.FindControl(xtpControlButton, ID_VIS_VISIONE).Visible = True
                End If
            End If
        Case 2
            If intAcc = ACC_NESSUNO Then
                If Not RibbonBar.FindControl(xtpControlButton, ID_VIS_TOTALI) Is Nothing Then RibbonBar.FindControl(xtpControlButton, ID_VIS_TOTALI).Visible = False
            ElseIf (intAcc > ACC_NESSUNO) And (intAcc <= ACC_TUTTI) Then
                If (intAcc And ACC_LETTURA) Then
                    If Not RibbonBar.FindControl(xtpControlButton, ID_VIS_TOTALI) Is Nothing Then RibbonBar.FindControl(xtpControlButton, ID_VIS_TOTALI).Enabled = True
                ElseIf (intAcc And ACC_MODIFICA) Then
                   If Not RibbonBar.FindControl(xtpControlButton, ID_VIS_TOTALI) Is Nothing Then RibbonBar.FindControl(xtpControlButton, ID_VIS_TOTALI).Visible = True
                End If
            End If
        Case 3
            If intAcc = ACC_NESSUNO Then
                Ling(SSCH2_VALORI).Visible = False
                Scheda(SSCH2_VALORI).Visible = False
            ElseIf (intAcc > ACC_NESSUNO) And (intAcc <= ACC_TUTTI) Then
                If (intAcc And ACC_LETTURA) Then
                    Ling(SSCH2_VALORI).Visible = True
                    Scheda(SSCH2_VALORI).Visible = True
                ElseIf (intAcc And ACC_MODIFICA) Then
                    Ling(SSCH2_VALORI).Visible = True
                    Scheda(SSCH2_VALORI).Visible = True
                End If
            End If
        Case 4
            If intAcc = ACC_NESSUNO Then
                Scheda(SSCH2_GRAFICO).Visible = False
                Ling(SSCH2_GRAFICO).Visible = False
            ElseIf (intAcc > ACC_NESSUNO) And (intAcc <= ACC_TUTTI) Then
                If (intAcc And ACC_LETTURA) Then
                    Ling(SSCH2_GRAFICO).Visible = True
                    Scheda(SSCH2_GRAFICO).Visible = True
                ElseIf (intAcc And ACC_MODIFICA) Then
                    Ling(SSCH2_GRAFICO).Visible = True
                    Scheda(SSCH2_GRAFICO).Visible = True
                End If
            End If
    End Select
End Sub

Private Function FiltroModificato(strCurFiltro As String) As Boolean
    FiltroModificato = False
    If cTraccia.pGestFiltroDati Then
        strCurFiltro = CFiltroDati.SQLFiltro
        FiltroModificato = (StrComp(mStrPrevFiltro, strCurFiltro, vbTextCompare) <> 0)
    End If
End Function

'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&
'           FUNZIONI PUBBLICHE DELLA FORM
'&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&

' NOME          : Visione
' DESCRIZIONE   : funzione di visione
' PARAMETRO 1   : nome della visione da rintracciare nel file INI
' PARAMETRO 2   : colonna di ordinamento
' PARAMETRO 3   : stringa SQL che indica il criterio WHERE per il primo livello
Public Function Visione(Optional strNomeVis As Variant, _
                    Optional strColOrdina As Variant, _
                    Optional strCriterio As Variant, _
                    Optional vntVariabili As Variant) As Boolean

Dim hWndVis As Long
Dim strTotale As String
Dim q As Integer

    Visione = True
On Error GoTo err_InitVisione
    'imposto valori default
    If (IsMissing(strNomeVis)) Then strNomeVis = mStrNomeVis
    If (IsMissing(strColOrdina)) Then strColOrdina = mStrOrdinamento
    If (IsMissing(strCriterio)) Then strCriterio = mStrCriterio
    
    q = InStr(strNomeVis, "|")
    If q > 0 Then
        strTotale = Mid(strNomeVis, q + 1)
        strNomeVis = Left(strNomeVis, q - 1)
    End If
    'imposto i parametri di visione
    mStrNomeVis = strNomeVis
    mStrOrdinamento = strColOrdina
    mStrCriterio = strCriterio
    mVntVariabili = vntVariabili
    'creo le classi di visione
    Set cTraccia = MXVI.CreaCTraccia
    Set cInterfaccia = MXVI.CreaCInterfaccia
On Error GoTo err_Visione
    'mostro la videata di visione
    If (LoadVisione()) Then
        
        'carico le impostazioni
        Dim sngSavedWidth As Single
        Dim sngSavedHeight As Single
        sngSavedWidth = Cast(MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MWForm" & MXNU.NTerminale & ".ini", "FRMVISIONI_" & Me.HelpContextID, "Width", "0"), vbSingle)
        sngSavedHeight = Cast(MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MWForm" & MXNU.NTerminale & ".ini", "FRMVISIONI_" & Me.HelpContextID, "Height", "0"), vbSingle)
        If (sngSavedWidth <> 0 And sngSavedWidth <> mSngMinWidth) Or (sngSavedHeight <> 0 And sngSavedHeight <> mSngMinHeight) Then
            Me.Move 0, 0, sngSavedWidth, sngSavedHeight
        End If

        'Inzializzazione Form per Metodo Evolus
        Call CambiaColoriControlli(Me)   'Da Mettere PRIMA di mResize
        '***SchedaTrovaBox.ShadowColor = SysGradientColor1
        On Local Error Resume Next
        '*****************************************************************
        'Set mResize = New MxResizer.ResizerEngine
        'If (Not mResize Is Nothing) Then
        '        Call mResize.Initialize(Me, , , , , 0, True, metodo, MXNU)
        'End If
        '*****************************************************************
        Call CentraFinestra(Me.hwnd)
        Call CambiaCharSet(Me)
        On Local Error GoTo 0
        
        If strTotale <> "" Then
            Call cTraccia.pLivello(1).GroupBy(strTotale, False)
            cTraccia.pLivello(1).NoGroupBy = True
            Call CommandBars_Execute(RibbonBar.FindControl(, ID_VIS_VISIONE))
            cTraccia.pLivello(1).NoGroupBy = False
        End If
        
        Call Me.Show
        
        'Se premuto Shift non passo in automatico sulla linguetta Visione
        'If mBolImpDefault And Not (GetAsyncKeyState(VK_LSHIFT) < 0 Or GetAsyncKeyState(VK_RSHIFT) < 0) Then Call Ling_GotFocus(SCH_VISIONE)
        If mBolImpDefault And Not (GetAsyncKeyState(VK_LSHIFT) < 0 Or GetAsyncKeyState(VK_RSHIFT) < 0) Then Call CommandBars_Execute(Me.CommandBars.FindControl(, ID_VIS_VISIONE))
    Else
        Call Unload(Me)
    End If

fine_Visione:
    On Local Error GoTo 0
    Exit Function

err_InitVisione:
    Visione = False
    Call MXNU.MsgBoxEX(1054, vbCritical, 1007, Array("Visione", Err.Number, Err.Description))
Resume fine_Visione

err_Visione:
    Visione = False
    Call MXNU.MsgBoxEX(1054, vbCritical, 1007, Array("Visione", Err.Number, Err.Description))
    Call Unload(Me)
Resume fine_Visione

    Resume
End Function

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
                cTraccia.Stampa (intSchOnTop1 = SCH_FILTRO), (intSchOnTop1 = SCH_VISIONE), (intSchOnTop1 = SCH_TOTALI), True
        Case MetFVisUtenteModifica
        Case MetFDettVisione
        Case MetFMostraCampiDBAnagr
        Case MetFPrimaZoom ' Rif. anomalia #6236
            If varparametro > 0 Then
                If Not (cTraccia Is Nothing) Then
                    If Not (cTraccia.pLivelloCorrente Is Nothing) Then
                        cTraccia.pLivelloCorrente.lngRigheVis = CLng(varparametro)
                        Call cTraccia.pLivelloCorrente.CaricaPagina(-1)
                    End If
                End If
            End If
        Case Else
    End Select

End Function

'Per Metodo Evolus
'Private Sub mResize_AfterResize()
'    Call AvvicinaLing(Me)
'End Sub

Private Sub CreateRibbon()
    Dim TabVis As RibbonTab, TabExp As RibbonTab
    
    Dim GroupSez As RibbonGroup
    Dim GroupLiv As RibbonGroup
    Dim GroupVis As RibbonGroup
    Dim GroupInfo As RibbonGroup
    Dim GroupAz As RibbonGroup
    Dim GroupExp As RibbonGroup
    Dim GroupTotali As RibbonGroup
    Dim Control As CommandBarControl
    Dim ControlPopUp As CommandBarPopup
    
    CommandBars.DeleteAll
    CommandBars.GlobalSettings.App = App
    CommandBars.EnableCustomization False
    'Risorse in lingua per i componenti Codejock
    CommandBars.GlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"
        
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched
    'RibbonBar.EnableFrameTheme
    
    Set TabVis = RibbonBar.InsertTab(ID_TABVIS_DATI, MXNU.CaricaStringaRes(203181))     '"&Dati"
    Set GroupSez = TabVis.Groups.AddGroup(MXNU.CaricaStringaRes(203206), ID_VISGROUP_GROUPSEZ) 'Sezione
    With GroupSez
        Set Control = .Add(xtpControlButton, ID_VIS_FILTRO, MXNU.CaricaStringaRes(21055))    '&Filtro
        Control.DescriptionText = MXNU.CaricaStringaRes(200435)   '"Visualizza la scheda Filtro (Ctrl+1)"
        Set Control = .Add(xtpControlButton, ID_VIS_VISIONE, MXNU.CaricaStringaRes(21056))   '&Visione
        Control.DescriptionText = MXNU.CaricaStringaRes(200436)   '"Visualizza la scheda Visione (Ctrl+2)"
        Set Control = .Add(xtpControlButton, ID_VIS_TOTALI, MXNU.CaricaStringaRes(21057))    '&Totali
        Control.DescriptionText = MXNU.CaricaStringaRes(200437)   '"Visualizza la scheda Totali (Ctrl+3)"
    End With
    
'    Set GroupLiv = TabVis.Groups.AddGroup(MXNU.CaricaStringaRes(203184), ID_VISGROUP_GROUPLIV)   '"Livelli"
'    With GroupLiv
'        .Add xtpControlButton, ID_VIS_INDIETRO, MXNU.CaricaStringaRes(54)
'        .Add xtpControlButton, ID_VIS_SELEZIONA, MXNU.CaricaStringaRes(55)
'        Set Control = .Add(xtpControlButton, ID_VIS_TERMINA, MXNU.CaricaStringaRes(57))
'        Control.Visible = False
'        .Add xtpControlButton, ID_VIS_RICARICA, MXNU.CaricaStringaRes(203189)  '"Ricarica Dati"
'    End With
    
    Set GroupVis = TabVis.Groups.AddGroup(MXNU.CaricaStringaRes(203185), ID_VISGROUP_GROUPVIS)   '"Visione"
    With GroupVis
        .Add xtpControlButton, ID_VIS_RICARICA, MXNU.CaricaStringaRes(203189)  '"Ricarica Dati"
        Set ControlPopUp = .Add(xtpControlSplitButtonPopup, ID_VIS_TIPOVIS, MXNU.CaricaStringaRes(203177))     '"Vis. Corrente"
        MOffSetTipoVis = 1
        Set ControlPopUp = .Add(xtpControlSplitButtonPopup, ID_VIS_OPZIONIVIS, MXNU.CaricaStringaRes(200438))  '"Opzioni Visione")
        MOffSetOpzioni = 1
        Set ControlPopUp = .Add(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA, MXNU.CaricaStringaRes(203176))   '"Totali&zzazioni Dinamiche"
            MOffSetRaggruppa = 1
            Set Control = ControlPopUp.CommandBar.Controls.Add(xtpControlButton, ID_VIS_RAGGRUPPA + MOffSetRaggruppa, MXNU.CaricaStringaRes(203190))   '"&Inserimento/Modifica"
            MOffSetRaggruppa = MOffSetRaggruppa + 1
            Set Control = ControlPopUp.CommandBar.Controls.Add(xtpControlButton, ID_VIS_RAGGRUPPA + MOffSetRaggruppa, MXNU.CaricaStringaRes(203191))   '"Nessun Raggruppamento"
            Control.BeginGroup = True
            MOffSetRaggruppa = MOffSetRaggruppa + 1
        .Add xtpControlButton, ID_VIS_RAGGRUPPACHIUDI, "Chiudi Totali"
        .Add xtpControlButton, ID_VIS_STAMPARAGGR, MXNU.CaricaStringaRes(203195)  '"Stampa"
        Set ControlPopUp = .Add(xtpControlSplitButtonPopup, ID_VIS_TROVA, MXNU.CaricaStringaRes(203203))   '"Trova"
            ControlPopUp.DescriptionText = MXNU.CaricaStringaRes(203196)   '"Trova testo o altro contenuto nella visione (CTRL+F)"
            Set Control = ControlPopUp.CommandBar.Controls.Add(xtpControlButton, ID_VIS_TROVASTD, MXNU.CaricaStringaRes(203197))  '"Standard"
            Control.DescriptionText = ""
            Set Control = ControlPopUp.CommandBar.Controls.Add(xtpControlButton, ID_VIS_TROVAAVZ, MXNU.CaricaStringaRes(203198))  '"Avanzato"
    End With
    
    Set GroupAz = TabVis.Groups.AddGroup(MXNU.CaricaStringaRes(203186), ID_VISGROUP_GROUPAZ)   '"Esegui Azione"
    With GroupAz
        Set ControlPopUp = .Add(xtpControlSplitButtonPopup, ID_VIS_AZIONI, MXNU.CaricaStringaRes(203178))    '"A&zioni"
        ControlPopUp.width = ControlPopUp.width + 70
        MOffSetAzioni = 1
    End With
    
    Set GroupTotali = TabVis.Groups.AddGroup(MXNU.CaricaStringaRes(203204), ID_VISGROUP_GROUPTOTALI)   '"Totali"
    With GroupTotali
        .Add xtpControlButton, ID_VIS_RICTOTALI, MXNU.CaricaStringaRes(40032)  '"Ricarica Totali"
    End With
    
    Set GroupInfo = TabVis.Groups.AddGroup(MXNU.CaricaStringaRes(203187), ID_VISGROUP_GROUPINFO)   '"Informazioni"
    With GroupInfo
        .Add xtpControlButton, ID_VIS_LEGENDA, MXNU.CaricaStringaRes(203179)  '"Legenda"
        .Add xtpControlButton, ID_VIS_INFO, MXNU.CaricaStringaRes(203180)  '"Informazioni Profilo"
    End With
    
    Set TabExp = RibbonBar.InsertTab(ID_TABVIS_EXPORT, MXNU.CaricaStringaRes(203182))   '"&Esporta"
    Set GroupExp = TabExp.Groups.AddGroup(MXNU.CaricaStringaRes(203188), ID_VISGROUP_GROUPEXP)   '"Export"
    With GroupExp
        .Add xtpControlButton, ID_VIS_EXPORTEXCEL, MXNU.CaricaStringaRes(40031)  '"Esporta su Excel"
        .Add xtpControlButton, ID_VIS_EXPORTOOCALC, MXNU.CaricaStringaRes(40055)  '"Esporta su OOCalc"
        Set ControlPopUp = .Add(xtpControlSplitButtonPopup, ID_VIS_EXPORTCRYSTAL, MXNU.CaricaStringaRes(40054))   '"Esporta su Crystal Reports"
            MOffSetExportCrystal = 1
            ControlPopUp.CommandBar.Controls.Add xtpControlButton, ID_VIS_EXPORTCRYSTAL + MOffSetExportCrystal, MXNU.CaricaStringaRes(40057)  '"&Crea Nuovo Report"
            MOffSetExportCrystal = MOffSetExportCrystal + 1
        Set ControlPopUp = .Add(xtpControlSplitButtonPopup, ID_VIS_EXPORTOLAP, MXNU.CaricaStringaRes(40056))   '"Esporta su OLAP"
            MOffSetExportOlap = 1
            ControlPopUp.CommandBar.Controls.Add xtpControlButton, ID_VIS_EXPORTOLAP + MOffSetExportOlap, MXNU.CaricaStringaRes(40058)  '"&Crea Modello di Analisi"
            MOffSetExportOlap = MOffSetExportOlap + 1
        .Add xtpControlButton, ID_VIS_EXPORTQLIK, MXNU.CaricaStringaRes(40059)  '"Esporta su QLikView"
    End With

    Set Control = RibbonBar.Controls.Add(xtpControlButton, ID_RIBBON_MINIMIZE, MXNU.CaricaStringaRes(203199))   '"Riduci a icona Barra Multifunzione (Ctrl+F1)"
    Control.DescriptionText = MXNU.CaricaStringaRes(203201)  '"Consente di visualizzare solo i nomi delle schede nella barra multifunzione"
    'Control.Height = 20
    Control.Style = xtpButtonAutomatic
    Control.flags = xtpFlagRightAlign
    
    Set Control = RibbonBar.Controls.Add(xtpControlButton, ID_RIBBON_EXPAND, MXNU.CaricaStringaRes(203200))     '"Espandi la barra multifunzione (Ctrl+F1)"
    Control.DescriptionText = MXNU.CaricaStringaRes(203202)   '"Visualizza la barra multifunzione in modo che rimanga sempre espansa"
    'Control.Height = 20
    Control.Style = xtpButtonAutomatic
    Control.flags = xtpFlagRightAlign

    'RibbonBar.GroupsHeight = 80
    RibbonBar.ContextMenuPresent = False
    
    'RibbonBar.AddSystemButton
    'RibbonBar.AllowQuickAccessCustomization = False
    'RibbonBar.ControlQuickAccess.Delete
    'RibbonBar.QuickAccessControls.DeleteAll
    RibbonBar.ShowQuickAccess = False
    RibbonBar.ShowGripper = False
    
    RibbonBar.AllowMinimize = True
    
    Me.CommandBars.AddImageList ImgListTB
    'CommandBars.Options.ShowTextBelowIcons = True
    If MXCtrl.TemaAttivo = "Office2010" Then
        Me.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
        Me.CommandBars.EnableOffice2007Frame True
        'CommandBars.VisualTheme = xtpThemeResource
        Me.CommandBars.VisualTheme = xtpThemeRibbon
    ElseIf MXCtrl.TemaAttivo = "Office2007" Then
        Me.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
        Me.CommandBars.EnableOffice2007Frame True
        'Me.CommandBars.VisualTheme = xtpThemeResource
        Me.CommandBars.VisualTheme = xtpThemeRibbon
    Else
        Me.CommandBars.VisualTheme = xtpThemeVisualStudio2008   'Anomalia nr. 12213
    End If
    RibbonBar.CommandBars.Options.KeyboardCuesUse = xtpKeyboardCuesUseAll
    Me.CommandBars.Options.KeyboardCuesShow = xtpKeyboardCuesShowAlways
    Me.CommandBars.Options.KeyboardCuesUse = xtpKeyboardCuesUseAll
    Me.CommandBars.KeyBindings.Add FCONTROL, vbKey1, ID_VIS_FILTRO
    Me.CommandBars.KeyBindings.Add FCONTROL, vbKey2, ID_VIS_VISIONE
    Me.CommandBars.KeyBindings.Add FCONTROL, vbKey3, ID_VIS_TOTALI
    Me.CommandBars.KeyBindings.Add FCONTROL, vbKeyF1, ID_RIBBON_MINIMIZE
    Me.CommandBars.KeyBindings.Add FCONTROL, vbKeyF1, ID_RIBBON_EXPAND
    Me.CommandBars.KeyBindings.Add FCONTROL, vbKeyF, ID_VIS_TROVASTD
    
    Dim ToolTipContext As ToolTipContext
    Set ToolTipContext = Me.CommandBars.ToolTipContext
    ToolTipContext.Style = xtpToolTipResource
    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
    ToolTipContext.ShowImage True, IMAGEBASE
    ToolTipContext.SetMargin 2, 2, 2, 2
    ToolTipContext.MaxTipWidth = 180
    
    Me.CommandBars.RecalcLayout
End Sub

Private Sub GestRaggruppa(ByVal ID As Long)
    If ID = ID_VIS_RAGGRUPPA + 1 Then   'Nuovo Raggruppamento
        Call cTraccia.pLivello(cTraccia.LivelloCorrente).GroupBy("", True)
        RibbonBar.FindControl(, ID_VIS_RAGGRUPPACHIUDI).Visible = True
        RibbonBar.FindControl(, ID_VIS_RAGGRUPPACHIUDI).Visible = True
    ElseIf ID = ID_VIS_RAGGRUPPA + 2 Then   'Nessun Raggruppamento - Torna alla visione
        With cTraccia
            If .pLivello(.LivelloCorrente).ssVisione.SheetCount > 1 Then
                Call .pLivello(.LivelloCorrente).GroupByClose
                Call .VisioneCaricaDati(.LivelloCorrente, .pLivelloCorrente.mIntVisione)
                Call GestRibbonBar(2)   'Per ripristinare correttamente la RibbonBar
                RibbonBar.FindControl(, ID_VIS_RAGGRUPPACHIUDI).Visible = False
            End If
        End With
    Else   'Attivo un Raggruppmento  esistente
        Dim ControlPopUp As CommandBarPopup
        Dim ControlItem As CommandBarControl
        Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA)
        Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, ID)
        If Not (ControlItem Is Nothing) Then
            If intSchOnTop1 = SCH_FILTRO Then
                Call cTraccia.pLivello(1).GroupBy(Trim(ControlItem.Caption), False)
                cTraccia.pLivello(1).NoGroupBy = True
                Call CommandBars_Execute(RibbonBar.FindControl(, ID_VIS_VISIONE))
                cTraccia.pLivello(1).NoGroupBy = False
            Else
                Call cTraccia.pLivello(cTraccia.LivelloCorrente).GroupBy(Trim(ControlItem.Caption), False)
                Call GestRibbonBar(1)
            End If
            RibbonBar.FindControl(, ID_VIS_RAGGRUPPACHIUDI).Visible = True
        End If
        Set ControlItem = Nothing
        Set ControlPopUp = Nothing
    End If
End Sub

'------------------------------------------------------------
'nome:          LoadBrioModels
'descrizione:   caricamento dei moduli brio
'ATTENZIONE: se si modifica questa funzione, ricordarsi di controllare anche la stessa funzione nella frmPopUp
'------------------------------------------------------------
Private Sub LoadBrioModels()
Dim oFSO As Scripting.FileSystemObject
Dim oFolder As Scripting.Folder
Dim oModel As Scripting.file
Dim intIdModel As Integer
Dim strPath As String
Dim strCommon As String
Dim ControlPopUp As CommandBarPopup
Dim ControlItem As CommandBarControl

    On Local Error Resume Next
    Set oFSO = New Scripting.FileSystemObject
    intIdModel = 0
    'modelli comuni
    strPath = MXNU.PercorsoPgm & "\BrioRepository\Common\" & cTraccia.pDefTraccia.pNomeTraccia
    strCommon = "common"
    GoSub LoadModelRibbonItems
    'modelli personali
    strPath = MXNU.PercorsoPgm & "\BrioRepository\" & MXNU.UtenteAttivo & "\" & cTraccia.pDefTraccia.pNomeTraccia
    strCommon = "personal"
    GoSub LoadModelRibbonItems

LoadBrioModels_END:
    Set oFSO = Nothing
    On Local Error GoTo 0
    Exit Sub
        
        
LoadModelRibbonItems:
    If (oFSO.FolderExists(strPath)) Then
        Set oFolder = oFSO.GetFolder(strPath)
        Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_EXPORTOLAP)
        For Each oModel In oFolder.Files
            intIdModel = intIdModel + 1
            If (StrComp(oFSO.GetExtensionName(oModel.path), "xml", vbTextCompare) = 0) Then
                'Load itemVisAzioniBrioModels(intIdModel)
                'itemVisAzioniBrioModels(intIdModel).Caption = oFSO.GetBaseName(oModel.path)
                'itemVisAzioniBrioModels(intIdModel).Tag = strCommon
                'itemVisAzioniBrioModels(intIdModel).Visible = True
                Set ControlItem = ControlPopUp.CommandBar.Controls.Add(xtpControlButton, ID_VIS_EXPORTOLAP + MOffSetExportOlap, oFSO.GetBaseName(oModel.path))
                MOffSetExportOlap = MOffSetExportOlap + 1
                ControlItem.Category = strCommon    'Manca la proprietà Tag: da verificare se questa proprietà ne fa le veci
            End If
        Next
    End If
Return
End Sub



'------------------------------------------------------------
'nome:          LoadCrystalReports
'descrizione:   caricamento dei report di Crystal
'parametri:
'ATTENZIONE: se si modifica questa funzione, ricordarsi di controllare anche la stessa funzione nella frmPopUp
'------------------------------------------------------------
Private Sub LoadCrystalReports()
Dim oFSO As Scripting.FileSystemObject
Dim oFolder As Scripting.Folder
Dim oReport As Scripting.file
Dim intIdReport As Integer
Dim strPath As String
Dim strCommon As String
Dim strTitle As String
Dim ControlPopUp As CommandBarPopup
Dim ControlItem As CommandBarControl

    On Local Error Resume Next
    Set oFSO = New Scripting.FileSystemObject
    intIdReport = 0
    'modelli comuni
    strPath = MXNU.PercorsoStampe & "\ExportVisioni\Common\" & cTraccia.pDefTraccia.pNomeTraccia
    strCommon = "common"
    GoSub LoadCrystalRibbonItems
    'modelli personali
    strPath = MXNU.PercorsoStampe & "\ExportVisioni\" & MXNU.UtenteAttivo & "\" & cTraccia.pDefTraccia.pNomeTraccia
    strCommon = "personal"
    GoSub LoadCrystalRibbonItems

LoadCrystalReports_END:
    Set oFSO = Nothing
    On Local Error GoTo 0
    Exit Sub
        
        
LoadCrystalRibbonItems:
    If (oFSO.FolderExists(strPath)) Then
        Set oFolder = oFSO.GetFolder(strPath)
        Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_EXPORTCRYSTAL)
        For Each oReport In oFolder.Files
            intIdReport = intIdReport + 1
            If (StrComp(oFSO.GetExtensionName(oReport.path), "rpt", vbTextCompare) = 0) Then
                Call MXCREP.LeggiReportInfo(oReport.path, strTitle, "")
                Set ControlItem = ControlPopUp.CommandBar.Controls.Add(xtpControlButton, ID_VIS_EXPORTCRYSTAL + MOffSetExportCrystal, strTitle)
                MOffSetExportCrystal = MOffSetExportCrystal + 1
                ControlItem.Category = strCommon & "\" & oFSO.GetBaseName(oReport.path)
                'Load itemVisAzioniCrystalReports(intIdReport)
                'itemVisAzioniCrystalReports(intIdReport).Caption = strTitle
                'itemVisAzioniCrystalReports(intIdReport).Tag = strCommon & "\" & oFSO.GetBaseName(oReport.path)
                'itemVisAzioniCrystalReports(intIdReport).Visible = True
            End If
        Next
    End If
Return
End Sub



'ATTENZIONE: se si modifica questa funzione, ricordarsi di controllare anche la stessa funzione nella frmPopUp
Private Sub LoadGroupBy(Optional bolClear As Boolean = False)
    Dim HSS As MXKit.CRecordSet
    Dim ControlPopUp As CommandBarPopup
    Dim ControlItem As CommandBarControl
    Dim intIDAzioni As Long
    Dim bolValido As Boolean
    
    On Local Error Resume Next
    
    If bolClear Then
        Dim i As Long
        Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA)
        For i = ID_VIS_RAGGRUPPA + 3 To ID_VIS_RAGGRUPPA + 20
            Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, i)
            If Not (ControlItem Is Nothing) Then
                ControlItem.Delete
            Else
                Exit For
            End If
        Next i
    End If
    
    If cTraccia.LivelloCorrente > 0 Then
        bolValido = (ssCellGetType(cTraccia.colLivelli(cTraccia.LivelloCorrente).ssVisione, 0, 1) <> SS_CELL_TYPE_CHECKBOX)
    Else
        bolValido = True
    End If
    'If ssCellGetType(cTraccia.colLivelli(cTraccia.LivelloCorrente).ssVisione, 0, 1) = SS_CELL_TYPE_CHECKBOX Then
    If Not bolValido Then
        CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA).Visible = False
        'itemVisAzioniFix(ITM_VIS_RAGGRUPPA).Visible = False
    Else
        'If itemVisAzioniRaggrs.Count = 1 Then
            Set HSS = MXDB.dbCreaRR(hndDBArchivi, "SELECT NomeImpostazione FROM TESTEIMPOSTAZIONIVISIONI WHERE NomeVisione=" & hndDBArchivi.FormatoSQL(cTraccia.pNomeTraccia, DB_TEXT) & " AND (NomeUtente='TRM' OR NomeUtente=" & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT) & ") ORDER BY NomeImpostazione")
            intIDAzioni = 1
            Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA)
            If bolClear Then MOffSetRaggruppa = 3
            While Not MXDB.dbFineTab(HSS)
                Set ControlItem = ControlPopUp.CommandBar.Controls.Add(xtpControlButton, ID_VIS_RAGGRUPPA + MOffSetRaggruppa, MXDB.dbGetCampo(HSS, NO_REPOSITION, "NomeImpostazione", ""))
                ControlItem.Category = CStr(intIDAzioni)
                MOffSetRaggruppa = MOffSetRaggruppa + 1
                'Load itemVisAzioniRaggrs(intIDAzioni)
                'itemVisAzioniRaggrs(intIDAzioni).Caption = MXDB.dbGetCampo(hSS, NO_REPOSITION, "NomeImpostazione", "")
                'itemVisAzioniRaggrs(intIDAzioni).Tag = intIDAzioni
                'itemVisAzioniRaggrs(intIDAzioni).Visible = True
                intIDAzioni = intIDAzioni + 1
                Call MXDB.dbSuccessivo(HSS)
            Wend
            Call MXDB.dbChiudiRR(HSS)
        'End If
    End If
    On Local Error GoTo 0
    
End Sub


Private Sub LoadAzioni()
    'modulo navigatore base -> carico sul menu azioni tutte le azioni disponibili
    If MXNU.ControlloModuliChiave(modNavigatoreBase) = 0 Then
        Dim Cazv As CAzioni
        Dim intIDAzioni As Integer
        Dim ControlPopUp As CommandBarPopup
        Dim ControlItem As CommandBarControl
        
        With cTraccia.colLivelli(cTraccia.LivelloCorrente)
            If .colAzioni.Count > 0 Then
                intIDAzioni = 1
                
                Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_AZIONI)
                On Local Error Resume Next
                ControlPopUp.CommandBar.Controls.DeleteAll
                MOffSetAzioni = 1
                For Each Cazv In .colAzioni
                    Set ControlItem = ControlPopUp.CommandBar.Controls.Add(xtpControlButton, ID_VIS_AZIONI + MOffSetAzioni, Cazv.Caption)
                    MOffSetAzioni = MOffSetAzioni + 1
                    ControlItem.Category = CStr(intIDAzioni)
                    intIDAzioni = intIDAzioni + 1
                Next
            Else
                RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = False
            End If
        End With
    Else
        RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = False
    End If

End Sub

Private Sub GestExportOlap(ByVal ID As Long)
    If (ID = ID_VIS_EXPORTOLAP + 1) Then
        If Not cTraccia.pLivello(cTraccia.LivelloCorrente).IsGroupBy Then
            Call cTraccia.Brio_CreateNewModel
        End If
    Else
        Dim ControlPopUp As CommandBarPopup
        Dim ControlItem As CommandBarControl
        
        Set ControlPopUp = CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_EXPORTOLAP)
        Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, ID)
        If Not (ControlItem Is Nothing) Then
            If Not cTraccia.pLivello(cTraccia.LivelloCorrente).IsGroupBy Then
                Call cTraccia.Brio_Export2Brio(ControlItem.Caption, (StrComp(ControlItem.Category, "common", vbTextCompare) = 0))
            End If
        End If
    End If

End Sub

Public Sub GestRibbonBar(ByVal NewSheet As Integer)
    Dim ControlPopUp As CommandBarPopup
    Dim ControlItem As CommandBarControl
    Dim i As Integer
    If NewSheet = 1 And cTraccia.colLivelli(cTraccia.LivelloCorrente).ssVisione.SheetCount = 1 Then NewSheet = 2
    Select Case NewSheet
        Case 1
            Set ControlPopUp = Me.CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA)
            For i = ID_VIS_RAGGRUPPA + 1 To ID_VIS_RAGGRUPPA + 20
                If (i - ID_VIS_RAGGRUPPA) <> 2 Then
                    Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, i)
                    If Not (ControlItem Is Nothing) Then
                        ControlItem.Enabled = True
                    End If
                End If
            Next i
            RibbonBar.FindControl(, ID_VIS_FILTRO).Visible = True
            RibbonBar.FindControl(, ID_VIS_TIPOVIS).Visible = False
            RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = (cTraccia.colLivelli(cTraccia.LivelloCorrente).CFiltro.colOpzioni.Count > 0)
            RibbonBar.FindControl(, ID_VIS_RICARICA).Visible = False
            RibbonBar.FindControl(, ID_VIS_STAMPARAGGR).Visible = True
            RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = True
            RibbonBar.FindControl(, ID_VIS_EXPORTCRYSTAL).Visible = False
            RibbonBar.FindControl(, ID_VIS_EXPORTOLAP).Visible = False
            RibbonBar.FindControl(, ID_VIS_EXPORTOOCALC).Visible = False
            RibbonBar.FindControl(, ID_VIS_EXPORTQLIK).Visible = False
            RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = False
            RibbonBar.FindControl(, ID_VIS_TROVA).Visible = False
            RibbonBar.FindGroup(ID_VISGROUP_GROUPTOTALI).Visible = False
        Case 2
            Set ControlPopUp = Me.CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA)
            For i = ID_VIS_RAGGRUPPA + 1 To ID_VIS_RAGGRUPPA + 20
                If (i - ID_VIS_RAGGRUPPA) <> 2 Then
                    Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, i)
                    If Not (ControlItem Is Nothing) Then
                        ControlItem.Enabled = True
                    End If
                End If
            Next i
            RibbonBar.FindControl(, ID_VIS_FILTRO).Visible = (cTraccia.pLivello(cTraccia.LivelloCorrente).ssVisione.SheetCount = 1)
            RibbonBar.FindControl(, ID_VIS_TIPOVIS).Visible = True
            If (cTraccia.colLivelli(cTraccia.LivelloCorrente).ssVisione.SheetCount = 1) Then
                RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = (cTraccia.colLivelli(cTraccia.LivelloCorrente).CFiltro.colOpzioni.Count > 0)
                RibbonBar.FindControl(, ID_VIS_TROVA).Visible = True
            Else
                RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = False
                RibbonBar.FindControl(, ID_VIS_TROVA).Visible = (cTraccia.colLivelli(cTraccia.LivelloCorrente).ssVisione.ActiveSheet = 2)
            End If
            RibbonBar.FindControl(, ID_VIS_RICARICA).Visible = (cTraccia.colLivelli(cTraccia.LivelloCorrente).ssVisione.SheetCount = 1)
            RibbonBar.FindControl(, ID_VIS_STAMPARAGGR).Visible = False
            RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = True
            RibbonBar.FindControl(, ID_VIS_EXPORTCRYSTAL).Visible = True
            RibbonBar.FindControl(, ID_VIS_EXPORTOLAP).Visible = True
            RibbonBar.FindControl(, ID_VIS_EXPORTOOCALC).Visible = True
            RibbonBar.FindControl(, ID_VIS_EXPORTQLIK).Visible = True
            RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = True
            RibbonBar.FindGroup(ID_VISGROUP_GROUPTOTALI).Visible = False
        Case 3
            RibbonBar.FindControl(, ID_VIS_FILTRO).Visible = False
            RibbonBar.Tab(ID_TABVIS_EXPORT).Visible = False
            RibbonBar.FindGroup(ID_VISGROUP_GROUPAZ).Visible = False
            RibbonBar.FindControl(, ID_VIS_TIPOVIS).Visible = False
            RibbonBar.FindControl(, ID_VIS_OPZIONIVIS).Visible = False
            RibbonBar.FindControl(, ID_VIS_RICARICA).Visible = False
            RibbonBar.FindControl(, ID_VIS_STAMPARAGGR).Visible = False
            RibbonBar.FindControl(, ID_VIS_TROVA).Visible = False
            RibbonBar.FindGroup(ID_VISGROUP_GROUPTOTALI).Visible = False
            Set ControlPopUp = Me.CommandBars.FindControl(xtpControlSplitButtonPopup, ID_VIS_RAGGRUPPA)
            For i = ID_VIS_RAGGRUPPA + 1 To ID_VIS_RAGGRUPPA + 20
                If (i - ID_VIS_RAGGRUPPA) <> 2 Then
                    Set ControlItem = ControlPopUp.CommandBar.Controls.Find(xtpControlButton, i)
                    If Not (ControlItem Is Nothing) Then
                        ControlItem.Enabled = False
                    End If
                End If
            Next i
    End Select

End Sub


Private Sub AttivaTrova(ByVal bolAvanzato As Boolean)
    SchedaTrovaBox.width = Frame(3).width + 200
    SchedaTrovaBox.Visible = True
    Frame(2).Left = 20000
    cTraccia.bolAvanzate = bolAvanzato
    SchedaTrovaBox.ZOrder 0
End Sub

