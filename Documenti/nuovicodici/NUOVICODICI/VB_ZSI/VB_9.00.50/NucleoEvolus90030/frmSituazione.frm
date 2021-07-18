VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{9060A1CD-0801-4EC9-839F-71695A0C5144}#1.0#0"; "MXKit.ocx"
Object = "{06022F8B-AE3E-45CE-9380-CADBE509AD34}#1.0#0"; "mxctrl.ocx"
Begin VB.Form frmSituazione 
   ClientHeight    =   6615
   ClientLeft      =   210
   ClientTop       =   1710
   ClientWidth     =   10620
   Icon            =   "frmSituazione.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   10620
   Begin MXCtrl.MWSchedaBox SchedaSit 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   11668
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
      ScaleWidth      =   10635
      ScaleHeight     =   6615
      Begin MXCtrl.MWSchedaBox SchedaSitArt 
         Height          =   5085
         Left            =   240
         TabIndex        =   28
         Top             =   1380
         Visible         =   0   'False
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   8969
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
         ScaleWidth      =   10125
         ScaleHeight     =   5085
         Begin VB.CheckBox ChkTot 
            Appearance      =   0  'Flat
            Caption         =   "Tot. Gen."
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
            Height          =   405
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   4560
            WhatsThisHelpID =   25088
            Width           =   1335
         End
         Begin VB.CheckBox Chk2UM 
            Appearance      =   0  'Flat
            Caption         =   "2 UM"
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
            Height          =   405
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   4560
            WhatsThisHelpID =   25087
            Width           =   1335
         End
         Begin VB.CheckBox Chk1UM 
            Appearance      =   0  'Flat
            Caption         =   "1 UM"
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
            Height          =   405
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   4560
            Value           =   1  'Checked
            WhatsThisHelpID =   25086
            Width           =   1335
         End
         Begin VB.CheckBox chkAllestCons 
            Appearance      =   0  'Flat
            Caption         =   "Allest./Cons."
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
            Height          =   405
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Frame FrmTot 
            Appearance      =   0  'Flat
            Caption         =   "Frame2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2865
            Left            =   630
            TabIndex        =   32
            Top             =   120
            Width           =   8805
            Begin FPSpreadADO.fpSpread Foglio 
               Height          =   2415
               Index           =   0
               Left            =   360
               TabIndex        =   33
               Top             =   270
               Width           =   8190
               _Version        =   524288
               _ExtentX        =   14446
               _ExtentY        =   4260
               _StockProps     =   64
               AutoCalc        =   0   'False
               BackColorStyle  =   1
               ColHeaderDisplay=   0
               DAutoCellTypes  =   0   'False
               DAutoFill       =   0   'False
               DAutoHeadings   =   0   'False
               DAutoSave       =   0   'False
               DAutoSizeCols   =   0
               EditEnterAction =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   2
               MaxRows         =   9
               NoBeep          =   -1  'True
               Protect         =   0   'False
               RestrictCols    =   -1  'True
               RowHeaderDisplay=   0
               ScrollBarExtMode=   -1  'True
               ScrollBars      =   0
               ShadowColor     =   12632256
               ShadowDark      =   8421504
               ShadowText      =   0
               SpreadDesigner  =   "frmSituazione.frx":0442
               UnitType        =   2
               UserResize      =   0
               VisibleCols     =   2
               VisibleRows     =   9
               AppearanceStyle =   0
            End
         End
         Begin VB.CommandButton CmdDettArt 
            Caption         =   "Dettaglio"
            Height          =   405
            Left            =   8160
            TabIndex        =   29
            Top             =   4560
            WhatsThisHelpID =   25043
            Width           =   1335
         End
         Begin VB.Frame fraAllestCons 
            Appearance      =   0  'Flat
            Caption         =   "Totali Allest./Cons."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1395
            Left            =   630
            TabIndex        =   34
            Top             =   3060
            Width           =   8805
            Begin FPSpreadADO.fpSpread ssAllestCons 
               Height          =   990
               Left            =   375
               TabIndex        =   35
               Top             =   300
               Width           =   8190
               _Version        =   524288
               _ExtentX        =   14446
               _ExtentY        =   1746
               _StockProps     =   64
               AutoCalc        =   0   'False
               BackColorStyle  =   1
               ColHeaderDisplay=   0
               DAutoCellTypes  =   0   'False
               DAutoFill       =   0   'False
               DAutoHeadings   =   0   'False
               DAutoSave       =   0   'False
               DAutoSizeCols   =   0
               DisplayColHeaders=   0   'False
               EditEnterAction =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   2
               MaxRows         =   4
               NoBeep          =   -1  'True
               Protect         =   0   'False
               RestrictCols    =   -1  'True
               RowHeaderDisplay=   0
               ScrollBarExtMode=   -1  'True
               ScrollBars      =   0
               ShadowColor     =   12632256
               ShadowDark      =   8421504
               ShadowText      =   0
               SpreadDesigner  =   "frmSituazione.frx":0A85
               UnitType        =   2
               UserResize      =   0
               VisibleCols     =   1
               VisibleRows     =   4
               AppearanceStyle =   0
            End
         End
         Begin VB.Frame FrameTotLifo 
            Appearance      =   0  'Flat
            Caption         =   "Totali Generali"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1395
            Left            =   630
            TabIndex        =   30
            Top             =   3075
            WhatsThisHelpID =   24759
            Width           =   8805
            Begin FPSpreadADO.fpSpread Foglio 
               Height          =   990
               Index           =   1
               Left            =   360
               TabIndex        =   31
               Top             =   270
               Width           =   5160
               _Version        =   524288
               _ExtentX        =   9102
               _ExtentY        =   1746
               _StockProps     =   64
               AutoCalc        =   0   'False
               BackColorStyle  =   1
               ColHeaderDisplay=   0
               DAutoCellTypes  =   0   'False
               DAutoFill       =   0   'False
               DAutoHeadings   =   0   'False
               DAutoSave       =   0   'False
               DAutoSizeCols   =   0
               DisplayColHeaders=   0   'False
               EditEnterAction =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   1
               MaxRows         =   4
               NoBeep          =   -1  'True
               Protect         =   0   'False
               RestrictCols    =   -1  'True
               RowHeaderDisplay=   0
               ScrollBarExtMode=   -1  'True
               ScrollBars      =   0
               ShadowColor     =   12632256
               ShadowDark      =   8421504
               ShadowText      =   0
               SpreadDesigner  =   "frmSituazione.frx":1038
               UnitType        =   2
               UserResize      =   1
               VisibleCols     =   1
               VisibleRows     =   4
               AppearanceStyle =   0
            End
         End
      End
      Begin MXCtrl.MWLinguetta LingSTot 
         Height          =   345
         Left            =   3360
         TabIndex        =   17
         Top             =   1380
         WhatsThisHelpID =   21057
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
      Begin VB.CommandButton Cmdretsit 
         Caption         =   "<<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9840
         TabIndex        =   3
         Top             =   120
         Width           =   435
      End
      Begin VB.TextBox txtb 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   0
         Left            =   3060
         TabIndex        =   2
         Top             =   135
         Width           =   6435
      End
      Begin VB.TextBox txtCodice 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   9120
         TabIndex        =   26
         Top             =   975
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.ComboBox cmbSituazione 
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
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   7245
      End
      Begin MXCtrl.MWLinguetta LingSFlt 
         Height          =   345
         Left            =   240
         TabIndex        =   9
         Top             =   1380
         WhatsThisHelpID =   21055
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
      Begin MXCtrl.MWLinguetta LingSVis 
         Height          =   345
         Left            =   1800
         TabIndex        =   13
         Top             =   1380
         WhatsThisHelpID =   21056
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
      Begin MXCtrl.XPToolButton tbsel 
         Height          =   290
         Index           =   0
         Left            =   9540
         Top             =   120
         Width           =   245
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
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   0
         Left            =   240
         Top             =   120
         Width           =   2715
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
         Caption         =   "etc_0"
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   2
         Left            =   240
         Top             =   960
         WhatsThisHelpID =   10520
         Width           =   2715
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
         Caption         =   "etc_2"
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   1
         Left            =   240
         Top             =   540
         Width           =   2715
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
         Caption         =   "etc_1"
      End
      Begin MXCtrl.MWEtichetta etcro 
         Height          =   300
         Index           =   1
         Left            =   3090
         Top             =   540
         Width           =   7215
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
      Begin MXCtrl.MWSchedaBox SchedaSVis 
         Height          =   4800
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   10120
         _ExtentX        =   17859
         _ExtentY        =   8467
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
         ScaleWidth      =   10125
         ScaleHeight     =   4800
         Begin MXCtrl.MWSchedaBox SchedaTrovaBox 
            Height          =   1155
            Left            =   60
            TabIndex        =   37
            Top             =   3600
            Width           =   10005
            _ExtentX        =   17648
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
            ScaleWidth      =   10005
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
               TabIndex        =   47
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
               Left            =   6480
               TabIndex        =   45
               Top             =   30
               Width           =   3435
               Begin FPSpreadADO.fpSpread ssOpzioni 
                  Height          =   735
                  Left            =   60
                  TabIndex        =   46
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   3315
                  _Version        =   524288
                  _ExtentX        =   5847
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
                  SpreadDesigner  =   "frmSituazione.frx":15B8
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
               TabIndex        =   38
               Top             =   30
               Width           =   6375
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
                  TabIndex        =   43
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
                  Left            =   5340
                  TabIndex        =   42
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
                  Left            =   5340
                  TabIndex        =   41
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
                  TabIndex        =   40
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
                  TabIndex        =   39
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   5115
               End
               Begin FPSpreadADO.fpSpread ssTrova 
                  Height          =   975
                  Left            =   120
                  TabIndex        =   44
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
                  SpreadDesigner  =   "frmSituazione.frx":2549
                  UserResize      =   0
                  VisibleCols     =   6
                  VisibleRows     =   3
                  AppearanceStyle =   0
               End
            End
         End
         Begin VB.ComboBox cmbVisione 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   30
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Frame FrameS 
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
            Height          =   3495
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   60
            Width           =   9855
            Begin FPSpreadADO.fpSpread ssVisione 
               DragIcon        =   "frmSituazione.frx":3D77
               Height          =   3120
               Left            =   120
               TabIndex        =   16
               Top             =   300
               Visible         =   0   'False
               Width           =   9615
               _Version        =   524288
               _ExtentX        =   16960
               _ExtentY        =   5503
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
               SpreadDesigner  =   "frmSituazione.frx":4081
               UnitType        =   2
               UserResize      =   0
               VirtualOverlap  =   15
               VirtualRows     =   15
               VisibleCols     =   10
               VisibleRows     =   6
               VScrollSpecial  =   -1  'True
               VScrollSpecialType=   1
               AppearanceStyle =   0
            End
         End
      End
      Begin MXCtrl.MWSchedaBox SchedaSFlt 
         Height          =   4800
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   10120
         _ExtentX        =   17859
         _ExtentY        =   8467
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
         ScaleWidth      =   10125
         ScaleHeight     =   4800
         Begin FPSpreadADO.fpSpread ssFiltroDati 
            Height          =   3855
            Left            =   150
            TabIndex        =   11
            Top             =   750
            Width           =   9795
            _Version        =   524288
            _ExtentX        =   17277
            _ExtentY        =   6800
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
            SpreadDesigner  =   "frmSituazione.frx":458E
            AppearanceStyle =   0
         End
         Begin MXKit.ctlImpostazioni CtlImpFiltro 
            Height          =   555
            Left            =   660
            TabIndex        =   12
            Top             =   120
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   979
         End
      End
      Begin MXCtrl.MWSchedaBox SchedaSTot 
         Height          =   4800
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   8467
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
         ScaleWidth      =   10095
         ScaleHeight     =   4800
         Begin MXCtrl.MWLinguetta LingSVal 
            Height          =   375
            Left            =   240
            TabIndex        =   19
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
         Begin MXCtrl.MWLinguetta LingSGrf 
            Height          =   375
            Left            =   1800
            TabIndex        =   24
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
         Begin MXCtrl.MWSchedaBox SchedaSVal 
            Height          =   4155
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   7329
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
            ScaleHeight     =   4155
            Begin VB.Frame FrameS 
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
               TabIndex        =   21
               Top             =   60
               WhatsThisHelpID =   24016
               Width           =   5475
               Begin VB.ComboBox cmbTotali 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   22
                  Top             =   240
                  Width           =   5235
               End
            End
            Begin FPSpreadADO.fpSpread ssTotali 
               DragIcon        =   "frmSituazione.frx":49E9
               Height          =   3180
               Left            =   120
               TabIndex        =   23
               Top             =   840
               Width           =   9375
               _Version        =   524288
               _ExtentX        =   16536
               _ExtentY        =   5609
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
               SpreadDesigner  =   "frmSituazione.frx":4CF3
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
         Begin MXCtrl.MWSchedaBox SchedaSGrf 
            Height          =   4155
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   7329
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
            ScaleHeight     =   4155
            Begin MXKit.CTLGrafico objChart 
               Height          =   4035
               Left            =   120
               TabIndex        =   27
               Top             =   60
               Width           =   9435
               _ExtentX        =   16642
               _ExtentY        =   7117
            End
         End
      End
   End
   Begin MXCtrl.MWLinguetta LingSit 
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Situazione"
   End
End
Attribute VB_Name = "frmSituazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

#If ISMETODO2005 Then
    'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1
#End If

'Rif. anomalia #8261
Public WithEvents MWAgt1 As MXKit.CAgenteAuto
Attribute MWAgt1.VB_VarHelpID = -1

Public Enum enmTipoSituazione
    SitArticolo = 0
    SitGenerica = 1
End Enum
Public TipoSituazione As enmTipoSituazione

'=================================
'   dichiarazione di costanti
'=================================
Const UMBASE = 1
Const UMSEC = 2
Const ROW_QTADISP = 9
Const ROW_GIACENZA = 6
'=================================
'   dichiarazione classi
'=================================
Private WithEvents mCSituazione As MXKit.cSituazione
Attribute mCSituazione.VB_VarHelpID = -1
Private WithEvents mCAnagrafica As MXKit.Anagrafica
Attribute mCAnagrafica.VB_VarHelpID = -1
'=================================
'   dichiarazione variabili
'=================================
Private Mboltestelab As Boolean
Private MstrArtElaborato As String
Private MbolArtElaborato As Boolean
Private MlngButtonMask As Long
Private MstrNomeSit As String
Private MvntCriterioAgg As Variant
Private MbolImpostaPrima As Boolean
Private mVntAData As Variant
Private mVntCodice As Variant
Private mstrCliente As String     ' rif.sch. A4562
Private mbolValidArticolo As Boolean     ' rif.sch. A4562
Private MbolCalcTot As Boolean
Private mbolUmSec As Boolean      ' rif.sch. A6537
Private mSngMinWidth As Single
Private mSngMinHeight As Single
Private mSngCurrentWidth As Single
Private mSngCurrentHeight As Single
Private TBSelHeight As Single
Private TBSelWidth As Single
Private OwnerForm As Form
Public InitOwnerForm As Boolean
Private cInterfaccia As MXKit.cInterfaccia

Private Sub Chk2UM_Click()

    ' inizio rif.sch. A6537
    mbolUmSec = False
    If Chk2UM.value = vbChecked Then
        mbolUmSec = True
    End If
    Call SpreadVisible
    Call CalcolaDatiArt(False)
    ' fine rif.sch. A6537

End Sub

' rif.sch. A6537
Private Sub SpreadVisible()

    If Chk2UM.value = vbChecked Then
        Call ssColShow(Foglio(0), 2)
    Else
        Call ssColHidden(Foglio(0), 2, True)
    End If
    If ChkTot.value = vbChecked Then
        FrameTotLifo.Visible = True
    Else
        FrameTotLifo.Visible = False
    End If
    If chkAllestCons.value = vbChecked Then
        fraAllestCons.Visible = True
        Chk2UM.Enabled = False
    Else
        Chk2UM.Enabled = True
        fraAllestCons.Visible = False
    End If

End Sub

'==========================================================================================
'                   EVENTI DELLA FORM
'==========================================================================================

Private Sub CalcolaAllest()

    Dim SQLAllest As String
    Dim SQLConsegne As String
    Dim hssAllest As CRecordSet
    Dim hssConsegne As CRecordSet
    Dim objArt As CVArt
    Dim strUM As String
    Dim vntValore As Variant
    Dim decQta As Variant
    Dim decQtaAllestita(1 To 2) As Variant, decQtaDaCons(1 To 2) As Variant
    Dim decDispNetta(1 To 2) As Variant, decGiacNetta(1 To 2) As Variant
    
    On Error GoTo ErrTrap
    SQLAllest = "SELECT * FROM STORICOALLESTIMENTI WHERE CODART =" & hndDBArchivi.FormatoSQL(mVntCodice, DB_TEXT)
    SQLConsegne = "SELECT * FROM STORICOCONSEGNE WHERE CODARTICOLO =" & hndDBArchivi.FormatoSQL(mVntCodice, DB_TEXT)
    
    ' crea gli oggetti
    Set hssAllest = MXDB.dbCreaSS(hndDBArchivi, SQLAllest, TIPO_SNAPSHOT)
    Set hssConsegne = MXDB.dbCreaSS(hndDBArchivi, SQLConsegne, TIPO_SNAPSHOT)
    Set objArt = MXART.CreaCVArt()
    objArt.Codice = CStr(mVntCodice)
    Call objArt.umLeggiUMPreferite(True)
    
    decQtaAllestita(UMBASE) = 0: decQtaAllestita(UMSEC) = 0
    ' esegue la conversione e la somma delle quantità allestite
    Do While Not MXDB.dbFineTab(hssAllest)
        strUM = MXDB.dbGetCampo(hssAllest, TIPO_SNAPSHOT, "UM", "")
        decQta = MXDB.dbGetCampo(hssAllest, TIPO_SNAPSHOT, "Quantita", 0)
        decQtaAllestita(UMBASE) = decQtaAllestita(UMBASE) + (objArt.umConvertiQuantita(decQta, strUM, objArt.umArticolo(UM_BASE)))
        decQtaAllestita(UMSEC) = decQtaAllestita(UMSEC) + (objArt.umConvertiQuantita(decQta, strUM, objArt.umArticolo(UM_SECONDARIA)))
        MXDB.dbSuccessivo hssAllest
    Loop
    
    decQtaDaCons(UMBASE) = 0: decQtaDaCons(UMSEC) = 0
    ' esegue la conversione e la somma delle quantità da consegnare
    Do While Not MXDB.dbFineTab(hssConsegne)
        strUM = MXDB.dbGetCampo(hssConsegne, TIPO_SNAPSHOT, "UM", "")
        decQta = MXDB.dbGetCampo(hssConsegne, TIPO_SNAPSHOT, "Quantita", 0)
        decQtaDaCons(UMBASE) = decQtaDaCons(UMBASE) + (objArt.umConvertiQuantita(decQta, strUM, objArt.umArticolo(UM_BASE)))
        decQtaDaCons(UMSEC) = decQtaDaCons(UMSEC) + (objArt.umConvertiQuantita(decQta, strUM, objArt.umArticolo(UM_SECONDARIA)))
        MXDB.dbSuccessivo hssConsegne
    Loop
    
    '------------------------------------------------------------------
    ' calcolo della disponibilità netta
    '------------------------------------------------------------------
    Foglio(0).GetText UMBASE, ROW_QTADISP, vntValore
    decDispNetta(UMBASE) = CDec(vntValore)
    decDispNetta(UMBASE) = decDispNetta(UMBASE) - (decQtaAllestita(UMBASE) + decQtaDaCons(UMBASE))
    Foglio(0).GetText UMSEC, ROW_QTADISP, vntValore
    decDispNetta(UMSEC) = CDec(vntValore)
    decDispNetta(UMSEC) = decDispNetta(UMSEC) - (decQtaAllestita(UMSEC) + decQtaDaCons(UMSEC))
    
    '------------------------------------------------------------------
    ' calcolo della giacenza netta
    '------------------------------------------------------------------
    Foglio(0).GetText UMBASE, ROW_GIACENZA, vntValore
    decGiacNetta(UMBASE) = CDec(vntValore)
    decGiacNetta(UMBASE) = decGiacNetta(UMBASE) - (decQtaAllestita(UMBASE) + decQtaDaCons(UMBASE))
    Foglio(0).GetText UMSEC, ROW_GIACENZA, vntValore
    decGiacNetta(UMSEC) = CDec(vntValore)
    decGiacNetta(UMSEC) = decGiacNetta(UMSEC) - (decQtaAllestita(UMSEC) + decQtaDaCons(UMSEC))
    
    '------------------------------------------------------------------
    ' scrive i risultati sul foglio e lo visualizza
    '------------------------------------------------------------------
    With ssAllestCons
        .SetText UMBASE, 1, Format(decQtaAllestita(UMBASE), MXNU.FORMATO_QUANTITA)
        .SetText UMBASE, 2, Format(decQtaDaCons(UMBASE), MXNU.FORMATO_QUANTITA)
        .SetText UMBASE, 3, Format(decDispNetta(UMBASE), MXNU.FORMATO_QUANTITA)
        .SetText UMBASE, 4, Format(decGiacNetta(UMBASE), MXNU.FORMATO_QUANTITA)
        .SetText UMSEC, 1, Format(decQtaAllestita(UMSEC), MXNU.FORMATO_QUANTITA)
        .SetText UMSEC, 2, Format(decQtaDaCons(UMSEC), MXNU.FORMATO_QUANTITA)
        .SetText UMSEC, 3, Format(decDispNetta(UMSEC), MXNU.FORMATO_QUANTITA)
        .SetText UMSEC, 4, Format(decGiacNetta(UMSEC), MXNU.FORMATO_QUANTITA)
    End With
    
    
FineSub:
    On Local Error Resume Next
    Call MXDB.dbChiudiSS(hssAllest)
    Call MXDB.dbChiudiSS(hssConsegne)
    Set hssAllest = Nothing
    Set hssConsegne = Nothing
    If Not (objArt Is Nothing) Then
        objArt.Termina
    End If
    Set objArt = Nothing
    On Local Error GoTo 0
    Exit Sub
    
ErrTrap:
Dim lngErrCod As Long
Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("Calcola allestimenti", lngErrCod, strErrDsc))
    Resume FineSub
    
End Sub

Private Sub chkAllestCons_Click()

    If chkAllestCons.value = vbChecked Then
    
        'rif.sch. A6537
        ' se l'unità di misura secondaria non è abilitata bisogna farlo
        If Not (mbolUmSec) Then
            Chk2UM.value = vbChecked
            Call Chk2UM_Click
        End If
        ChkTot.Enabled = True
        ChkTot.value = vbUnchecked
        Call CalcolaAllest
        chkAllestCons.Enabled = False
    End If
    Call SpreadVisible
   
End Sub

Private Sub ChkTot_Click()
    If ChkTot.value = vbChecked Then
        chkAllestCons.value = vbUnchecked   'rif.sch. A6537
        chkAllestCons.Enabled = True
        ChkTot.Enabled = False
        MbolCalcTot = True
        Call CalcolaTotArt   'rif.sch. A6537
    End If
    Call SpreadVisible   'rif.sch. A6537
End Sub

Private Sub CmdDettArt_Click()
    Me.KeyPreview = True
    If Mboltestelab Then
        Call CreaSituazione
    End If
    'Cmdretsit.Visible = True
    Cmdretsit.Enabled = True
    SchedaSitArt.Visible = False
    SchedaSitArt.ZOrder vbSendToBack
    Me.KeyPreview = True
End Sub

Private Sub Cmdretsit_Click()
    Me.KeyPreview = False
    'Cmdretsit.Visible = False
    Cmdretsit.Enabled = False
    SchedaSitArt.Visible = True
    SchedaSitArt.ZOrder
End Sub

Private Sub Foglio_ColWidthChange(Index As Integer, ByVal Col1 As Long, ByVal Col2 As Long)
    If Index = 0 Then
        ssAllestCons.ColWidth(1) = Foglio(0).ColWidth(1)
        ssAllestCons.ColWidth(2) = Foglio(0).ColWidth(2)
    End If
End Sub

'rif.sch. A4562
Private Sub Form_Initialize()
     mbolValidArticolo = False
     mstrCliente = ""
End Sub

Private Sub Form_Activate()
    Call MXNU.Attiva_Toolbar(hwnd, MlngButtonMask)
    Call MXNU.ImpostaFormAttiva(Me)
End Sub

Private Sub Form_Load()
    Set OwnerForm = MXNU.DammiFormAttiva
    Mboltestelab = True
    MlngButtonMask = BTN_TUTTI_MASK
    Me.Caption = MXNU.CaricaStringaRes(23164)
    Call CentraFinestra(hwnd)
    Call MXNU.LeggiRisorseControlli(Me)
    'Rif. anomalia #8261
    Set MWAgt1 = MXAA.CreaCAgenteAuto()
    
    mSngMinHeight = SchedaSit.Height
    mSngMinWidth = SchedaSit.width
    mSngCurrentHeight = mSngMinHeight
    mSngCurrentWidth = mSngMinWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MWAgt1 = Nothing
    Set mCAnagrafica = Nothing   ' rif.sch. A4744
    Set cInterfaccia = Nothing
    Set mCSituazione = Nothing
    Set OwnerForm = Nothing
    Set frmSituazione = Nothing
End Sub

'metto in primo piano la form delle situazioni
Private Sub BringToFront()
    If Me.Visible Then
        Call Me.ZOrder(vbBringToFront)
        Me.WindowState = vbNormal
    End If
End Sub

'==========================================================================================
'                   PROPRIETA' PUBBLICHE DELLA FORM
'==========================================================================================
Public Property Get pCodice() As Variant
    pCodice = mVntCodice
End Property

Public Property Let pCodice(new_Valore As Variant)
    If new_Valore <> mVntCodice Then
        mVntCodice = new_Valore
        Screen.MousePointer = vbHourglass
        Call mCAnagrafica.AssegnaCampo(mCAnagrafica.NomeControlloCodice, mVntCodice)
        If Not (mCSituazione Is Nothing) Then
            If mCSituazione.SituazioneImpostata Then
                mCSituazione.strCriterio = mVntCodice
            End If
        End If
        If TipoSituazione = SitArticolo Then
            Call DefSpread
            If Not (mbolUmSec) Then
                Call ssColHidden(Foglio(0), 2, True)
            End If
            Call CalcolaDatiArtRefresh      'rif.sch. A6537
        End If
        Screen.MousePointer = vbDefault
    End If
    Call BringToFront
End Property

' rif.sch. A4562 - Aggiunta la proprietà pCliente
Public Property Get pCliente() As String
    pCliente = mstrCliente
End Property

Public Property Let pCliente(new_Valore As String)
    mstrCliente = new_Valore
End Property

Public Property Get pValidArticolo() As Boolean
    pValidArticolo = mbolValidArticolo
End Property

Public Property Let pValidArticolo(new_Valore As Boolean)
    mbolValidArticolo = new_Valore
End Property

Public Property Get pNomeSituazione() As String
    pNomeSituazione = MstrNomeSit
End Property

'==========================================================================================
'                   FUNZIONI PUBBLICHE DELLA FORM
'==========================================================================================
Public Function Situazione(strNomeSit As String, _
    Optional vntCodice As Variant, _
    Optional vntCriterioAgg As Variant, _
    Optional bolImpostaPrima As Boolean = False) As Boolean

On Error GoTo err_Situazione
    Situazione = True
    Screen.MousePointer = vbHourglass
    'inizializzo l'anagrafica
    Set mCAnagrafica = MXVA.CreaCAnagrafica(strNomeSit, Me)
    Call mCAnagrafica.Disegna
    
    MstrNomeSit = strNomeSit
    MbolImpostaPrima = bolImpostaPrima
    If Not IsMissing(vntCodice) Then
        pCodice = vntCodice
    End If
    If Not IsMissing(vntCriterioAgg) Then
        MvntCriterioAgg = vntCriterioAgg
    End If
    If TipoSituazione = SitGenerica Then
        Call CreaSituazione
    ElseIf TipoSituazione = SitArticolo Then
        mbolUmSec = False  'rif.sch. A6537
        SchedaSitArt.Visible = True
        SchedaSitArt.ZOrder 0
        ' Rif. anomalia #8261 assegnamento ID fisso per situazione articolo
        Me.HelpContextID = 908
    End If
    ' Rif. anomalia #8261
    Call MWAgt1.RegistraAgenteFrm(Me)
#If ISMETODO2005 Then
    'Inzializzazione Form per Metodo Evolus
    Call CambiaColoriControlli(Me)   'Da Mettere PRIMA di mResize
    SchedaTrovaBox.ShadowColor = SysGradientColor1
    On Local Error Resume Next
    Set mResize = New MxResizer.ResizerEngine
    If (Not mResize Is Nothing) Then
        Call mResize.Initialize(Me, , , , , 0, True, metodo, MXNU)
    End If
#End If
    Call CentraFinestra(Me.hwnd)
    Call CambiaCharSet(Me)
    On Local Error GoTo 0
    Me.Show
    
fine_Situazione:
    Screen.MousePointer = vbDefault
    On Local Error GoTo 0
    Exit Function

err_Situazione:
    Situazione = False
    Call MXNU.MsgBoxEX(1054, vbCritical, 1007, Array("Visione", Err.Number, Err.Description))
Resume fine_Situazione
Resume

End Function

Private Sub LingSit_GotFocus()
    If Not (mCSituazione Is Nothing) Then Call mCSituazione.LingSit_GotFocus
End Sub

Private Sub LingSTot_GotFocus()
    If Not (mCSituazione Is Nothing) Then Call mCSituazione.LingTot_GotFocus
End Sub

Private Sub mCAnagrafica_DopoValidazioneGruppo(ByVal strNomeGruppo As String, vntNewValore As Variant, ByVal enmTipoEvento As MXKit.SetTipoValidazione)
    'inizio rif.sch.831 - cambio codice -> cambio situazione
    If strNomeGruppo = mCAnagrafica.NomeControlloCodice Then
    
        Call AggSituazione(vntNewValore)   ' rif.sch. A5351
        
        'nasconde il foglio delle qtà allestite e da cons.
        ' fraAllestCons.Visible = False        ' rif.sch. A6537
        ' chkAllestCons.Value = vbUnchecked     ' rif.sch. A6537
    End If
    'fine rif.sch.831
End Sub

Private Sub mCSituazione_PrimaCaricamentoRiga(ByVal lngRow As Long, HrsVis As MXKit.CRecordSet, bolSuccesso As Boolean)
    If StrComp(mCSituazione.CMyTraccia.pNomeTraccia, "VIS_ANALISI_DISP", vbTextCompare) = 0 And lngRow = 2 Then   'Anomlia nr. 12297
        If mCSituazione.CMyTraccia.pLivelloCorrente.lngRigheCar = 1 Then
            HrsVis.recSet.MoveFirst
        End If
    End If
End Sub

Private Sub mCSituazione_PrimaEsecuzioneQuery(strQuery As String, bolSuccesso As Boolean)

    With mCSituazione
        If StrComp(.CMyTraccia.pNomeTraccia, "VIS_ANALISI_DISP", vbTextCompare) = 0 Then
            Dim vntDataElab As Variant
            Dim strWHERE As String
            vntDataElab = .CMyFiltroDati.ParAgg("DataElab").ValoreFormula
            'chiamata alla routine che esegue l'analisi e riempie la tabella temporanea
            'nota:  la strQuery della visione viene reimpostata assegnando alla parte "WHERE"
            '       solo l'IDSessione corrente!
            'RIF.A#9620 - compongo la clausola where
            strWHERE = .CMyTraccia.pLivelloCorrente.strFltSit & " AND " & .CMyFiltroDati.SQLFiltro
            bolSuccesso = CreaTempPerAnalisiDisp(strQuery, _
                            strWHERE, _
                            .CMyTraccia.pLivelloCorrente.SQLDammiORDERBY(.CMyTraccia.pLivelloCorrente.mIntVisione), _
                            vntDataElab)
        End If
    End With

End Sub

Private Sub mCSituazione_RecordSetTotaliVuoto(HrsTot As MXKit.CRecordSet, ByVal intTotImposta As Integer, bolRichiediParziali As Boolean)
    
    If TipoSituazione = SitArticolo Then
        Call Totali_AddRecordIniziali(mCSituazione.CMyTraccia, HrsTot, bolRichiediParziali, cmbTotali.listIndex + 1, 2, 5, ssFiltroDati, True)
    End If
    
End Sub

Private Sub mCSituazione_RichiediRecordSetTotali(HrsTot As MXKit.CRecordSet, bolSuccesso As Boolean)
    
    Select Case UCase(mCSituazione.CMyTraccia.pNomeTraccia)
        Case "VIS_MOVMAG", "VIS_MOVMAG_BASE"
            Call Totali_AddRecordIniziali(mCSituazione.CMyTraccia, HrsTot, True, cmbTotali.listIndex + 1, 2, 5, ssFiltroDati, True)
    End Select

End Sub

Private Sub mCSituazione_ValidazionePers(ByVal strNomeValid As String, ByVal strNomeCmpValid As String, bolEseguiValStd As Boolean, vntNewValore As Variant)
    Call ValidPersFiltri(strNomeValid, strNomeCmpValid, bolEseguiValStd, vntNewValore)
End Sub

Private Sub SchedaSFlt_Paint()
    Call SchedaOmbreggiaControlli(SchedaSFlt)
End Sub

Private Sub SchedaSGrf_Paint()
    Call SchedaOmbreggiaControlli(SchedaSGrf)
End Sub

Private Sub SchedaSit_Paint()
    Call SchedaOmbreggiaControlli(SchedaSit)
End Sub

Private Sub SchedaSitArt_Paint()
    Call SchedaOmbreggiaControlli(SchedaSitArt)
    FrmTot.Caption = MXNU.CaricaStringaRes(24814)
End Sub

Private Sub SchedaSTot_Paint()
    Call SchedaOmbreggiaControlli(SchedaSTot)
End Sub

Private Sub SchedaSVal_Paint()
    Call SchedaOmbreggiaControlli(SchedaSVal)
End Sub

Private Sub SchedaSVis_Paint()
    Call SchedaOmbreggiaControlli(SchedaSVis)
End Sub

Private Sub tbSel_Click(Index As Integer)

    ' -- modifica per Rif. sch. 3654 e 3777 il 10/07/2002
    If TipoSituazione = SitArticolo Then
        Call SelezionaArticolo
    Else
        Call mCAnagrafica.TBselClick("tbSel_" & Index, "")
    End If
End Sub

'rif.sch. A5148
Private Sub CalcolaDatiArtRefresh()

    ' inizio rif.sch. A6537
    Call CalcolaDatiArt(True)
    If MbolCalcTot Then
        Call CalcolaTotArt
    End If
    If chkAllestCons.value = vbChecked Then
        Call CalcolaAllest
    End If
    Call SpreadVisible
    ' fine rif.sch. A6537
    
End Sub

Private Sub CalcolaTotArt()

#If ISNUCLEO = 0 Then
Dim CValArt As CValArticolo
    
    On Local Error GoTo ERR_CalcolaTotArt
    ' FrameTotLifo.Visible = True
    Set CValArt = MXART.CreaCValArticolo
    Call CValArt.CalcolaDatiCaricoScarico(True, mVntCodice)
    Call Foglio(1).SetText(1, 1, Format(CValArt.TotQtaCarico, MXNU.FORMATO_QUANTITA))
    'Anomalia nr. 4905
    If MXNU.UsaEuro Then
        Call Foglio(1).SetText(1, 2, Format(CValArt.TotValCaricoEuro, MXNU.FORMATO_EURO_TOTALE))
    Else
        Call Foglio(1).SetText(1, 2, Format(CValArt.TotValCarico, MXNU.FORMATO_LIRE_TOTALE))
    End If
       
    Call CValArt.CalcolaDatiCaricoScarico(False, mVntCodice)
    Call Foglio(1).SetText(1, 3, Format(CValArt.TotQtaScarico, MXNU.FORMATO_QUANTITA))
    If MXNU.UsaEuro Then
        Call Foglio(1).SetText(1, 4, Format(CValArt.TotValScaricoEuro, MXNU.FORMATO_EURO_TOTALE))
    Else
        Call Foglio(1).SetText(1, 4, Format(CValArt.TotValScarico, MXNU.FORMATO_LIRE_TOTALE))
    End If
        
END_CalcolaTotArt:
    On Local Error Resume Next
    Set CValArt = Nothing
    On Local Error GoTo 0
    Exit Sub
        
ERR_CalcolaTotArt:
Dim lngErrCod As Long
Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("Calcola Totale articolo", lngErrCod, strErrDsc))
    Resume END_CalcolaTotArt
    
#End If

End Sub

Private Sub CalcolaDatiArt(ByVal bol1Um As Boolean)
    
#If ISNUCLEO = 0 Then
    Dim CDispArt As MXBusiness.cDispArticolo
    
    On Local Error GoTo ERR_CalcolaDatiArt
    Set CDispArt = MXART.CreaCDispArticolo()
    
    'Anomalia nr. 4851
    CDispArt.DepositiDisponibili = True
    If bol1Um Then
        Call CDispArt.CalcolaDisp(mVntCodice, MXNU.AnnoAttivo, False, "", MAG_TUTTE_LE_UBICAZIONI, MAG_TUTTE_LE_PARTITE, False, False, MXNU.DataIniMag, AData)
        Call Foglio(0).SetText(1, 1, Format(CDispArt.GiacenzaIniziale, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(1, 2, Format(CDispArt.Carichi, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(1, 3, Format(CDispArt.Scarichi, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(1, 4, Format(CDispArt.ResiDaCarico, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(1, 5, Format(CDispArt.ResidaScarico, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(1, 6, Format(CDispArt.Giacenza, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(1, 7, Format(CDispArt.Ordinato, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(1, 8, Format(CDispArt.Impegnato, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(1, 9, Format(CDispArt.Disponibilita, MXNU.FORMATO_QUANTITA))
    End If
    
    If mbolUmSec Then
        Call CDispArt.CalcolaDisp(mVntCodice, MXNU.AnnoAttivo, True, "", MAG_TUTTE_LE_UBICAZIONI, MAG_TUTTE_LE_PARTITE, False, False, MXNU.DataIniMag, AData)
        Call Foglio(0).SetText(2, 1, Format(CDispArt.GiacenzaIniziale, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(2, 2, Format(CDispArt.Carichi, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(2, 3, Format(CDispArt.Scarichi, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(2, 4, Format(CDispArt.ResiDaCarico, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(2, 5, Format(CDispArt.ResidaScarico, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(2, 6, Format(CDispArt.Giacenza, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(2, 7, Format(CDispArt.Ordinato, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(2, 8, Format(CDispArt.Impegnato, MXNU.FORMATO_QUANTITA))
        Call Foglio(0).SetText(2, 9, Format(CDispArt.Disponibilita, MXNU.FORMATO_QUANTITA))
    End If
    
END_CalcolaDatiArt:
    On Local Error Resume Next
    Set CDispArt = Nothing
    On Local Error GoTo 0
    Exit Sub
    
ERR_CalcolaDatiArt:
Dim lngErrCod As Long
Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("Calcola Dati articolo", lngErrCod, strErrDsc))
    Resume END_CalcolaDatiArt
    
#End If
    
End Sub

Private Sub DefSpread()
    Dim strUM1 As String
    Dim strUM2 As String
    Dim hSS As MXKit.CRecordSet
    Dim intq As Integer

    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT UM FROM ARTICOLIUMPREFERITE WHERE TipoUM=1 AND CodArt=" & hndDBArchivi.FormatoSQL(mVntCodice, DB_TEXT))
    If Not MXDB.dbFineTab(hSS) Then
        strUM1 = MXDB.dbGetCampo(hSS, NO_REPOSITION, "UM", "")
    End If
    intq = MXDB.dbChiudiSS(hSS)
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT UM FROM ARTICOLIUMPREFERITE WHERE TipoUM=2 AND CodArt=" & hndDBArchivi.FormatoSQL(mVntCodice, DB_TEXT))
    If Not MXDB.dbFineTab(hSS) Then
        strUM2 = MXDB.dbGetCampo(hSS, NO_REPOSITION, "UM", "")
    End If
    intq = MXDB.dbChiudiSS(hSS)
    Call Foglio(0).SetText(1, 0, MXNU.CaricaStringaRes(30707, strUM1))
    Call Foglio(0).SetText(2, 0, MXNU.CaricaStringaRes(30708, strUM2))

    Call ssSpreadImposta(Foglio(0))
    Call MXNU.ssCaricaCaptionColonne(Foglio(0), 0, Foglio(0).MaxCols, 0, Foglio(0).MaxRows)

    Call ssSpreadImposta(Foglio(1))
    Call MXNU.ssCaricaCaptionColonne(Foglio(1), 0, Foglio(1).MaxCols, 0, Foglio(1).MaxRows)

    ' imposta il foglio delle quantità allestite e ad consegnare
    Call ssSpreadImposta(ssAllestCons)
    Call MXNU.ssCaricaCaptionColonne(ssAllestCons, 0, ssAllestCons.MaxCols, 0, ssAllestCons.MaxRows)
    ssAllestCons.ColWidth(1) = Foglio(0).ColWidth(1)
    ssAllestCons.ColWidth(2) = Foglio(0).ColWidth(2)
    Call ssDefText(ssAllestCons, 1, 1, 30, , , SS_CELL_H_ALIGN_RIGHT)
    Call ssDefText(ssAllestCons, 1, 2, 30, , , SS_CELL_H_ALIGN_RIGHT)
    Call ssDefText(ssAllestCons, 1, 3, 30, , , SS_CELL_H_ALIGN_RIGHT)
    Call ssDefText(ssAllestCons, 1, 4, 30, , , SS_CELL_H_ALIGN_RIGHT)
    Call ssDefText(ssAllestCons, 2, 1, 30, , , SS_CELL_H_ALIGN_RIGHT)
    Call ssDefText(ssAllestCons, 2, 2, 30, , , SS_CELL_H_ALIGN_RIGHT)
    Call ssDefText(ssAllestCons, 2, 3, 30, , , SS_CELL_H_ALIGN_RIGHT)
    Call ssDefText(ssAllestCons, 2, 4, 30, , , SS_CELL_H_ALIGN_RIGHT)
    fraAllestCons.Visible = False
    
End Sub

Sub CreaSituazione()
    Dim bolSituazione As Boolean

    Call MXNU.MostraMsgInfo(24701)
    'inizializzo la situazione
    Set cInterfaccia = MXVI.CreaCInterfaccia
    Set mCSituazione = MXVI.CreaCSituazione()
    With mCSituazione
        .pImpostaControlli = False
        bolSituazione = .Inizializza(IIf(InitOwnerForm, OwnerForm, Me), txtCodice, _
            cmbSituazione, _
            LingSit, SchedaSit, _
            LingSFlt, SchedaSFlt, _
            LingSVis, SchedaSVis, _
            LingSTot, SchedaSTot, _
            LingSVal, SchedaSVal, _
            LingSGrf, SchedaSGrf, _
            ssFiltroDati, _
            TxtTrova, txtAvanzato, ssTrova, _
            comAvanzato, comTrova, _
            ssOpzioni, lstHelp, _
            cmbTotali, ssTotali, _
            ssVisione, cmbVisione, _
            objChart, _
            CtlImpFiltro, _
            MstrNomeSit)

        'RIF. A#6560 - Imposta l'HelpID della form caricandolo dalla situazione
        Me.HelpContextID = .pHelpId

        If bolSituazione Then
            'imposto il criterio aggiuntivo
            If MvntCriterioAgg <> "" Then
                .pCriterioAgg = MvntCriterioAgg
            End If
            If .SituazioneImpostata Then
                .strCriterio = mVntCodice
            End If
            'imposto la prima situazione disponibile
            On Local Error Resume Next
            cmbSituazione.listIndex = 0
            'Anomalia nr. 6895
            'If MbolImpostaPrima Then
            '    Call mCSituazione.LingTot_GotFocus
            '    cmbTotali.ListIndex = 0
            'End If
            On Local Error GoTo 0
        Else
            Unload Me
        End If
    End With
    Call MXNU.MostraMsgInfo("")
    Mboltestelab = False
End Sub

Public Function AzioniMetodo(setAzione As enmFunzioniMetodo, Optional varparametro As Variant) As Variant
    
    Select Case setAzione
        Case MetFInserisci
        Case MetFRegistra
        Case MetFAnnulla
        Case MetFPrecedente: Call MetPrecedente
        Case MetFSuccessivo: Call MetSuccessivo
        Case MetFPrimo:      Call MetPrimo
        Case MetFUltimo:     Call MetUltimo
        Case MetFStampa
        If SchedaSitArt.Visible = False Then
            Call mCSituazione.StampaSituazione
        End If
    End Select
    
End Function

Sub MetPrimo()

    Call mCAnagrafica.FrecceSpostamento(BTN_PRIMO)
    If TipoSituazione = SitArticolo Then
        mVntCodice = mCAnagrafica.grinput(mCAnagrafica.NomeControlloCodice).ValoreCorrente
        Call CalcolaDatiArtRefresh        'rif.sch. A5148
    End If

End Sub

Sub MetPrecedente()

    Call mCAnagrafica.FrecceSpostamento(BTN_PREC)
    If TipoSituazione = SitArticolo Then
        mVntCodice = mCAnagrafica.grinput(mCAnagrafica.NomeControlloCodice).ValoreCorrente
        Call CalcolaDatiArtRefresh        'rif.sch. A5148
    End If

End Sub

Sub MetSuccessivo()

    Call mCAnagrafica.FrecceSpostamento(BTN_SUCC)
    If TipoSituazione = SitArticolo Then
        mVntCodice = mCAnagrafica.grinput(mCAnagrafica.NomeControlloCodice).ValoreCorrente
        Call CalcolaDatiArtRefresh        'rif.sch. A5148
    End If

End Sub

Sub MetUltimo()

    Call mCAnagrafica.FrecceSpostamento(BTN_ULTIMO)
    If TipoSituazione = SitArticolo Then
        mVntCodice = mCAnagrafica.grinput(mCAnagrafica.NomeControlloCodice).ValoreCorrente
        Call CalcolaDatiArtRefresh        'rif.sch. A5148
    End If

End Sub

Public Property Let AData(new_Valore As Variant)

    mVntAData = new_Valore

End Property

Public Property Get AData() As Variant

    If Not IsDate(mVntAData) Then
        AData = MXNU.DataFineMag
    Else
        AData = mVntAData
    End If

End Property

Private Sub txtb_GotFocus(Index As Integer)
    'selezione del contenuto
    Call SelContenuto(txtb(Index))
End Sub

Private Sub txtb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim intInd As Integer
Dim strRifCtrl As String
    'gestione ctrl+s per selezione
    If (KeyCode = vbKeyS And Shift = vbCtrlMask) Then
        strRifCtrl = mCAnagrafica.NomeControllo2NomeVariabile("txtb_" & Index)
        intInd = mCAnagrafica.IndiceToolButton(strRifCtrl)
        If intInd >= 0 Then Call tbSel_Click(intInd)
    End If

End Sub

Private Sub txtb_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strRifCtrl As String
    'maschera tasti non validi
    strRifCtrl = mCAnagrafica.NomeControllo2NomeVariabile("txtb_" & Index)
    Call CtrlKey(KeyAscii, mCAnagrafica.grinput(strRifCtrl).TipoInput)

End Sub

Private Sub txtb_LostFocus(Index As Integer)
Dim strRifCtrl As String
    'validazione + assegnamento campo
    Screen.MousePointer = vbHourglass
    strRifCtrl = mCAnagrafica.NomeControllo2NomeVariabile("txtb_" & Index)
    If (mCAnagrafica.grinput(strRifCtrl).ValoreCorrente <> txtb(Index).text) Then
        If (strRifCtrl = mCAnagrafica.NomeControlloCodice) Then
            'campo codice -> richiesta salvataggio
            'If (TestSalva() <> tsritorna) Then
                '----------------
                ' -- modifica per Rif. sch. 3654 e 3777 il 10/07/2002
                If (TipoSituazione = SitArticolo) Or (mbolValidArticolo) Then  'rif.sch. A4562
                    Call ValidaArticolo(txtb(0).text)
                Else
                    Call mCAnagrafica.AssegnaCampo(strRifCtrl, txtb(Index).text)
                End If
                '----------------
                'Call ImpostaScheda(mIntSchOnTop)
            'Else
            '    txtb(Index).Text = mCAnagrafica.GrInput(strRifCtrl).ValoreCorrente
            'End If
        Else
            'campo non codice -> assegnamento
            Call mCAnagrafica.AssegnaCampo(strRifCtrl, txtb(Index).text)
        End If
    End If
    Screen.MousePointer = vbDefault
    
End Sub

' ------------------------------------------
' Inizio Rif. sch. 3654 e 3777
' Funzioni inserite per validare gli articoli varianti
' ------------------------------------------

Private Sub ValidaArticolo(strCodice As String)

Dim xCodArt As MXBusiness.CVArt

    Set xCodArt = MXART.CreaCVArt()
    xCodArt.Codice = strCodice
    If xCodArt.Valida(CHIEDIVAR_TUTTE, False, , 0, False, , mstrCliente) Then   'rif.sch. A4562
        txtb(0).text = xCodArt.Codice
        etcro(1).Caption = xCodArt.Descrizione
        mCAnagrafica.grinput(mCAnagrafica.NomeControlloCodice).ValoreCorrente = xCodArt.Codice
        mVntCodice = xCodArt.Codice
        Call CalcolaDatiArtRefresh      'rif.sch. A6537
        Call AggSituazione(mVntCodice)  ' rif.sch. A5351
    Else
        Call MXNU.MsgBoxEX(1848, vbExclamation, 1007)
    End If
    If Not (xCodArt Is Nothing) Then xCodArt.Termina
    Set xCodArt = Nothing

End Sub

Private Sub SelezionaArticolo()

Dim colCampiRit As Collection

    Set colCampiRit = New Collection
    If mCAnagrafica.GrInput_Selezione(mCAnagrafica.NomeControlloCodice, colCampiRit) Then
        Call ValidaArticolo(colCampiRit(mCAnagrafica.NomeControlloCodice))
    End If
    Set colCampiRit = Nothing
                    
End Sub

' ------------------------------------------
' Fine Rif. sch. 3654 e 3777
' funzioni inserite per validare le varianti
' ------------------------------------------

' rif.sch. A5351
' Non aggiornava il dettaglio se si premeva il pulsante di selezione all'interno del dettaglio
Sub AggSituazione(ByVal strCodArt As String)

    If Len(strCodArt) > 0 Then
        If Not (mCSituazione Is Nothing) Then
            If mCSituazione.SituazioneImpostata Then
                Screen.MousePointer = vbHourglass
                mVntCodice = strCodArt
                mCSituazione.strCriterio = mVntCodice
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
    
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
#If ISMETODO2005 Then
    'Per Metodo Evolus
    If Not Cancel Then
        On Local Error Resume Next
        If (Not mResize Is Nothing) Then
                mResize.Terminate
                Set mResize = Nothing
        End If
        On Local Error GoTo 0
    End If
#End If
End Sub

#If ISMETODO2005 Then
    'Per Metodo Evolus
    Private Sub mResize_AfterResize()
        Dim sngWidth As Single
        Dim sngHeight As Single
        Dim sngDeltaWidth As Single
        Dim sngDeltaHeight As Single
            
        On Local Error Resume Next
        Call AvvicinaLing(Me)
        
        sngWidth = SchedaSit.width
        sngDeltaWidth = sngWidth - mSngCurrentWidth
        mSngCurrentWidth = sngWidth
    
        sngHeight = SchedaSit.Height
        sngDeltaHeight = sngHeight - mSngCurrentHeight
        mSngCurrentHeight = sngHeight
    
        SchedaSVis.Height = SchedaSVis.Height + sngDeltaHeight
        SchedaSVis.width = SchedaSVis.width + sngDeltaWidth
        SchedaSVal.Height = SchedaSVal.Height + sngDeltaHeight
        SchedaSVal.width = SchedaSVal.width + sngDeltaWidth
        SchedaSTot.Height = SchedaSTot.Height + sngDeltaHeight
        SchedaSTot.width = SchedaSTot.width + sngDeltaWidth
        SchedaSFlt.Height = SchedaSFlt.Height + sngDeltaHeight
        SchedaSFlt.width = SchedaSFlt.width + sngDeltaWidth
            
        ssFiltroDati.Height = ssFiltroDati.Height + sngDeltaHeight
        ssFiltroDati.width = ssFiltroDati.width + sngDeltaWidth
        
        ssTotali.Height = ssTotali.Height + sngDeltaHeight
        ssTotali.width = ssTotali.width + sngDeltaWidth
        
        ssVisione.Height = ssVisione.Height + sngDeltaHeight
        ssVisione.width = ssVisione.width + sngDeltaWidth
        cmbVisione.Left = cmbVisione.Left + sngDeltaWidth
        
        If MXNU.ResizeProporzionale Then
            If (Me.Height - mSngMinHeight) > (mSngMinHeight * MXNU.PercResizeProporzionale / 100) Then ssSpreadSetFontSize ssVisione, 10 Else ssSpreadSetFontSize ssVisione, 8
        End If
        
        FrameS(0).Height = FrameS(0).Height + sngDeltaHeight
        FrameS(0).width = FrameS(0).width + sngDeltaWidth
        
        SchedaTrovaBox.Top = SchedaTrovaBox.Top + sngDeltaHeight '- 50
        SchedaTrovaBox.width = SchedaTrovaBox.width + sngDeltaWidth
            
        SchedaSitArt.Top = LingSFlt.Top
        SchedaSitArt.Left = LingSFlt.Left
        SchedaSitArt.Height = SchedaSVis.Height + LingSFlt.Height
        SchedaSitArt.width = SchedaSVis.width
        If Not (mCSituazione Is Nothing) Then
            Call mCSituazione.CMyTraccia.CalcVisibleRows
        End If
        On Local Error GoTo 0
    End Sub
#End If

Private Sub LingSFlt_GotFocus()
    If Not (mCSituazione Is Nothing) Then Call mCSituazione.LingFlt_GotFocus
End Sub

Private Sub LingSGrf_GotFocus()
    If Not (mCSituazione Is Nothing) Then Call mCSituazione.LingGrf_GotFocus
End Sub

Private Sub LingSVal_GotFocus()
    If Not (mCSituazione Is Nothing) Then Call mCSituazione.LingVal_GotFocus
End Sub

Private Sub LingSVis_GotFocus()
    If Not (mCSituazione Is Nothing) Then Call mCSituazione.LingVis_GotFocus
End Sub

#If ISMETODO2005 Then
Private Sub mResize_AfterInitialize()
    Dim objCtrl As Control

    With mResize
        Call .ResizableControls.RemoveControl(SchedaSVis)
        Call .ResizableControls.RemoveControl(SchedaSFlt)
        Call .ResizableControls.RemoveControl(SchedaSVal)
        Call .ResizableControls.RemoveControl(SchedaSitArt)
        Call .ResizableControls.RemoveControl(SchedaSTot)
        Call .ResizableControls.RemoveControl(SchedaTrovaBox)
        Call .ResizableControls.RemoveControl(ssVisione)
        Call .ResizableControls.RemoveControl(ssTrova)
        Call .ResizableControls.RemoveControl(TxtTrova)
        Call .ResizableControls.RemoveControl(txtAvanzato)
        Call .ResizableControls.RemoveControl(cmbVisione)
        Call .ResizableControls.RemoveControl(FrameS(0))
        Call .ResizableControls.RemoveControl(FrameS(4))
        Call .ResizableControls.RemoveControl(Frame(2))
        Call .ResizableControls.RemoveControl(Frame(3))
        Call .ResizableControls.RemoveControl(comAvanzato)
        Call .ResizableControls.RemoveControl(comTrova)
        Call .ResizableControls.RemoveControl(ssOpzioni)
        Call .ResizableControls.RemoveControl(LingSFlt)
        Call .ResizableControls.RemoveControl(LingSVis)
        Call .ResizableControls.RemoveControl(LingSTot)
        Call .ResizableControls.RemoveControl(LingSVal)
        Call .ResizableControls.RemoveControl(LingSGrf)
        Call .ResizableControls.RemoveControl(cmbTotali)
        Call .ResizableControls.RemoveControl(ssTotali)
        Call .ResizableControls.RemoveControl(ssFiltroDati)
        Call .ResizableControls.RemoveControl(CtlImpFiltro)
        For Each objCtrl In CtlImpFiltro.Controls
            Call .ResizableControls.RemoveControl(objCtrl)
        Next
        Call .ResizableControls.RemoveControl(cmbSituazione)
        Call .ResizableControls.RemoveControl(etc(0))
        Call .ResizableControls.RemoveControl(etc(1))
        Call .ResizableControls.RemoveControl(etc(2))
        Call .ResizableControls.RemoveControl(etcro(1))
        Call .ResizableControls.RemoveControl(txtb(0))
        Call .ResizableControls.RemoveControl(tbsel(0))
        Call .ResizableControls.RemoveControl(Cmdretsit)
    End With
End Sub
#End If

Private Sub ssVisione_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Call cInterfaccia.VisioneButtonClicked(mCSituazione.CMyTraccia, 1, Col, Row, ButtonDown)
End Sub

Private Sub ssVisione_KeyPress(KeyAscii As Integer)
    Call cInterfaccia.VisioneKeyPress(mCSituazione.CMyTraccia, 1, KeyAscii)
End Sub

Private Sub ssVisione_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call cInterfaccia.VisioneLeaveCell(mCSituazione.CMyTraccia, 1, Col, Row, NewCol, NewRow, Cancel)
End Sub

Private Sub ssVisione_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    Call cInterfaccia.VisioneSheetChanged(mCSituazione.CMyTraccia, 1, OldSheet, NewSheet)
End Sub

Private Sub ssVisione_SheetChanging(ByVal OldSheet As Integer, ByVal NewSheet As Integer, Cancel As Variant)
    Call cInterfaccia.VisioneSheetChanging(mCSituazione.CMyTraccia, 1, OldSheet, NewSheet, Cancel)
End Sub

Private Sub ssVisione_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    Call cInterfaccia.VisioneTopLeftChange(mCSituazione.CMyTraccia, 1, OldLeft, OldTop, NewLeft, NewTop)
End Sub


