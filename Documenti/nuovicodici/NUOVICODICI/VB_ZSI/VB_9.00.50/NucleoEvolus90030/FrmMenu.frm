VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CODEJOCK.COMMANDBARS.V15.3.1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "CODEJOCK.SHORTCUTBAR.V15.3.1.OCX"
Object = "{6F0498EE-AF09-42C0-B462-CAB975BC79BD}#1.0#0"; "mxctrl.ocx"
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "frmMenu"
   ClientHeight    =   8370
   ClientLeft      =   4095
   ClientTop       =   1845
   ClientWidth     =   4725
   LinkTopic       =   "frmMenu"
   ScaleHeight     =   8370
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImgLstModuli 
      Left            =   1200
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   65
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":0000
            Key             =   "Contabilita"
            Object.Tag             =   "301"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":051B
            Key             =   "Cicli"
            Object.Tag             =   "302"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":0A6F
            Key             =   "Automazione"
            Object.Tag             =   "305"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":0FCE
            Key             =   "Enasarco"
            Object.Tag             =   "306"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":151F
            Key             =   "Metodo98"
            Object.Tag             =   "307"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":21F9
            Key             =   "Stampa"
            Object.Tag             =   "308"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2720
            Key             =   "CartAp"
            Object.Tag             =   "309"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2C04
            Key             =   "CartCh"
            Object.Tag             =   "310"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":30BD
            Key             =   "Stampe"
            Object.Tag             =   "311"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":363F
            Key             =   "Testo"
            Object.Tag             =   "312"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":3B6F
            Key             =   "help"
            Object.Tag             =   "313"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":40E3
            Key             =   "Exe"
            Object.Tag             =   "314"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":47DD
            Key             =   "Magazzino"
            Object.Tag             =   "333"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":4CDD
            Key             =   "Archivi"
            Object.Tag             =   "318"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":5222
            Key             =   "Documenti"
            Object.Tag             =   "328"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":5758
            Key             =   "VBanco"
            Object.Tag             =   "319"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":5C99
            Key             =   "Visione"
            Object.Tag             =   "317"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":6393
            Key             =   "Schedulatore"
            Object.Tag             =   "321"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":692F
            Key             =   "project"
            Object.Tag             =   "355"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":7029
            Key             =   "E98Commerce"
            Object.Tag             =   "323"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":75C4
            Key             =   "Pianificazione"
            Object.Tag             =   "340"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":7B35
            Key             =   "Distinta"
            Object.Tag             =   "329"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":8065
            Key             =   "Quality"
            Object.Tag             =   "332"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":83DD
            Key             =   "Navigatore"
            Object.Tag             =   "335"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":8923
            Key             =   "AnaBilancio"
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":8E17
            Key             =   "Risorse"
            Object.Tag             =   "326"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":9335
            Key             =   "Commesse"
            Object.Tag             =   "327"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":9699
            Key             =   "Costi"
            Object.Tag             =   "324"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":9BDA
            Key             =   "Consegne"
            Object.Tag             =   "322"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":9F43
            Key             =   "keyprog"
            Object.Tag             =   "339"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":A474
            Key             =   "Utilita"
            Object.Tag             =   "316"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":A9BE
            Key             =   "TargetAgenti"
            Object.Tag             =   "346"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":ACC3
            Key             =   "congestione"
            Object.Tag             =   "356"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":AFE8
            Key             =   "Ritenute"
            Object.Tag             =   "334"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":B4C0
            Key             =   "Strumenti"
            Object.Tag             =   "315"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":B6AF
            Key             =   "packing"
            Object.Tag             =   "345"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":B9EE
            Key             =   "Analisi"
            Object.Tag             =   "336"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":BF0F
            Key             =   "wizard"
            Object.Tag             =   "354"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":C3A9
            Key             =   "ImpExp"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":C93F
            Key             =   "ImportExport"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":CED5
            Key             =   "EsportazioneDati"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":D46B
            Key             =   "Produzione"
            Object.Tag             =   "338"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":D9F5
            Key             =   "Cespiti"
            Object.Tag             =   "331"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":DF72
            Key             =   "Negozi"
            Object.Tag             =   "330"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":E4B3
            Key             =   "OLAPITEM"
            Object.Tag             =   "351"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":E9E8
            Key             =   "OLAPCART"
            Object.Tag             =   "352"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":EEB8
            Key             =   "AgentiDO"
            Object.Tag             =   "350"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":F20E
            Key             =   "OrdineCliente"
            Object.Tag             =   "347"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":F76B
            Key             =   "OrdineForn"
            Object.Tag             =   "348"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":FCC8
            Key             =   "MOLAP"
            Object.Tag             =   "363"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":101DC
            Key             =   "menunew"
            Object.Tag             =   "344"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":10648
            Key             =   "menuold"
            Object.Tag             =   "343"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":10B28
            Key             =   "Doc"
            Object.Tag             =   "337"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":11085
            Key             =   "SearchNode"
            Object.Tag             =   "361"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":112E4
            Key             =   "Rename"
            Object.Tag             =   "359"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":116FB
            Key             =   "SearchFilter"
            Object.Tag             =   "360"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":11952
            Key             =   "DelGrp"
            Object.Tag             =   "358"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":11AC9
            Key             =   "TOOLS"
            Object.Tag             =   "325"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":11FF4
            Key             =   "Del"
            Object.Tag             =   "357"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":12254
            Key             =   "exportdati"
            Object.Tag             =   "364"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1294E
            Key             =   "costoprodotto"
            Object.Tag             =   "365"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":13048
            Key             =   "tesoreria"
            Object.Tag             =   "366"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":13742
            Key             =   "contratti"
            Object.Tag             =   "367"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":13E3C
            Key             =   "gruppiacquisto"
            Object.Tag             =   "368"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":14536
            Key             =   "Aiot"
            Object.Tag             =   "369"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox bottomPanel 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   5345
      Width           =   4695
      Begin MSComctlLib.TreeView TwPreferiti 
         DragIcon        =   "FrmMenu.frx":14C30
         Height          =   2775
         Left            =   0
         TabIndex        =   5
         Top             =   285
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4895
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImgLstModuli"
         Appearance      =   0
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4695
         _Version        =   983043
         _ExtentX        =   8281
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Preferiti"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox upPanel 
      BorderStyle     =   0  'None
      Height          =   4570
      Left            =   0
      ScaleHeight     =   4575
      ScaleWidth      =   4695
      TabIndex        =   1
      Top             =   660
      Width           =   4695
      Begin MSComctlLib.TreeView TrVSearch 
         DragIcon        =   "FrmMenu.frx":15072
         Height          =   3645
         Left            =   720
         TabIndex        =   7
         Top             =   660
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   6429
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImgLstModuli"
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
      Begin MSComctlLib.TreeView TrwModuli 
         DragIcon        =   "FrmMenu.frx":154B4
         Height          =   5115
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   9022
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImgLstModuli"
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
   Begin MXCtrl.MWSplitter MWSplitter1 
      Height          =   120
      Left            =   0
      TabIndex        =   0
      Top             =   5230
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   212
      UpControls      =   "upPanel"
      BottomControls  =   "bottomPanel"
      RightControls   =   ""
      LeftControls    =   ""
      MouseIcon       =   "FrmMenu.frx":158F6
      Orientation     =   1
   End
   Begin MSComctlLib.ImageList ImgLstModuli16x16 
      Left            =   1860
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   68
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":15A58
            Key             =   "Contabilita"
            Object.Tag             =   "301"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":15FF2
            Key             =   "Cicli"
            Object.Tag             =   "302"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":168CC
            Key             =   "IMPEXP"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":16E66
            Key             =   "Automazione"
            Object.Tag             =   "305"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":17740
            Key             =   "Enasarco"
            Object.Tag             =   "306"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":17CDA
            Key             =   "Metodo98"
            Object.Tag             =   "307"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":189B4
            Key             =   "Stampa"
            Object.Tag             =   "308"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":18F4E
            Key             =   "CartAp"
            Object.Tag             =   "309"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":194E8
            Key             =   "CartCh"
            Object.Tag             =   "310"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":19A82
            Key             =   "Stampe"
            Object.Tag             =   "311"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1A01C
            Key             =   "Testo"
            Object.Tag             =   "312"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1A5B6
            Key             =   "help"
            Object.Tag             =   "313"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1AB50
            Key             =   "Exe"
            Object.Tag             =   "314"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1B0EA
            Key             =   "Strumenti"
            Object.Tag             =   "315"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1B684
            Key             =   "Utilita"
            Object.Tag             =   "316"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1BC1E
            Key             =   "Visione"
            Object.Tag             =   "317"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1C1B8
            Key             =   "Archivi"
            Object.Tag             =   "318"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1C752
            Key             =   "VBanco"
            Object.Tag             =   "319"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1CCEC
            Key             =   "Analisi"
            Object.Tag             =   "336"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1D286
            Key             =   "AnaBilancio"
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1D820
            Key             =   "Schedulatore"
            Object.Tag             =   "321"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1DDBA
            Key             =   "Consegne"
            Object.Tag             =   "322"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1E354
            Key             =   "E98Commerce"
            Object.Tag             =   "323"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1E8EE
            Key             =   "Costi"
            Object.Tag             =   "324"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1EE88
            Key             =   "TOOLS"
            Object.Tag             =   "325"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1FCDA
            Key             =   "Risorse"
            Object.Tag             =   "326"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":20274
            Key             =   "Commesse"
            Object.Tag             =   "327"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2080E
            Key             =   "Documenti"
            Object.Tag             =   "328"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":20DA8
            Key             =   "Distinta"
            Object.Tag             =   "329"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":21342
            Key             =   "Negozi"
            Object.Tag             =   "330"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":218DC
            Key             =   "Cespiti"
            Object.Tag             =   "331"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":21E76
            Key             =   "Quality"
            Object.Tag             =   "332"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":22410
            Key             =   "Magazzino"
            Object.Tag             =   "333"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":229AA
            Key             =   "Ritenute"
            Object.Tag             =   "334"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":22F44
            Key             =   "Navigatore"
            Object.Tag             =   "335"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":234DE
            Key             =   "Doc"
            Object.Tag             =   "337"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":24330
            Key             =   "Produzione"
            Object.Tag             =   "338"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":248CA
            Key             =   "keyprog"
            Object.Tag             =   "339"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":24E64
            Key             =   "Pianificazione"
            Object.Tag             =   "340"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":25CB6
            Key             =   "UNLOCK"
            Object.Tag             =   "341"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":26250
            Key             =   "LOCK"
            Object.Tag             =   "342"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":267EA
            Key             =   "menuold"
            Object.Tag             =   "343"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":26D84
            Key             =   "menunew"
            Object.Tag             =   "344"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2E08E
            Key             =   "packing"
            Object.Tag             =   "345"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2E628
            Key             =   "TargetAgenti"
            Object.Tag             =   "346"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2EBC2
            Key             =   "OrdineCliente"
            Object.Tag             =   "347"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2F15C
            Key             =   "OrdineForn"
            Object.Tag             =   "348"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2F6F6
            Key             =   "ImportExport"
            Object.Tag             =   "349"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2FC90
            Key             =   "AgentiDO"
            Object.Tag             =   "350"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":3022A
            Key             =   "OLAPITEM"
            Object.Tag             =   "351"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":307C4
            Key             =   "OLAPCART"
            Object.Tag             =   "352"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":30D5E
            Key             =   "MOLAP"
            Object.Tag             =   "363"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":312F8
            Key             =   "wizard"
            Object.Tag             =   "354"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":31892
            Key             =   "project"
            Object.Tag             =   "355"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":31E2C
            Key             =   "congestione"
            Object.Tag             =   "356"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":323C6
            Key             =   "Del"
            Object.Tag             =   "357"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":32960
            Key             =   "DelGrp"
            Object.Tag             =   "358"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":32EFA
            Key             =   "Rename"
            Object.Tag             =   "359"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":33494
            Key             =   "SearchFilter"
            Object.Tag             =   "360"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":33A2E
            Key             =   "SearchNode"
            Object.Tag             =   "361"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":33FC8
            Key             =   "RunProg"
            Object.Tag             =   "362"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":34562
            Key             =   "Preferiti"
            Object.Tag             =   "363"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":34AFC
            Key             =   "exportdati"
            Object.Tag             =   "364"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":35096
            Key             =   "costoprodotto"
            Object.Tag             =   "365"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":35630
            Key             =   "tesoreria"
            Object.Tag             =   "366"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":35BCA
            Key             =   "contratti"
            Object.Tag             =   "367"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":36164
            Key             =   "gruppiacquisto"
            Object.Tag             =   "368"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":366FE
            Key             =   "Aiot"
            Object.Tag             =   "369"
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      _Version        =   983043
      _ExtentX        =   8281
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "Menu Completo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Image ImgDrag 
      Height          =   480
      Index           =   1
      Left            =   4980
      Picture         =   "FrmMenu.frx":36C98
      Top             =   1320
      Width           =   480
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   0
      Top             =   0
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FSHIFT = 4
Const FCONTROL = 8
Const FALT = 16


Const ID_POPUP_ADDGROUP = 2001
Const ID_POPUP_RENAME = 2002
Const ID_POPUP_DELGROUP = 2003
Const ID_POPUP_DELITEM = 2004
Const ID_POPUP_RUNPROG = 2005
Const ID_POPUP_ADDPREFERITI = 2006

Const ID_TLBSEARCH_TEXT = 3001
Const ID_TLBSEARCH_FILTER = 3002
Const ID_TLBSEARCH_NODESEARCH = 3003

Const VK_F2 = &H71
Const VK_DELETE = &H2E

'Dim McolMenu As CcolMenu
'Dim MlngMenuID As Long

Dim MSepIdx As Long
Dim MOffSetIconID As Long
Private WithEvents ModShortcutBar As ShortcutBar
Attribute ModShortcutBar.VB_VarHelpID = -1
Dim NodoDraggato As MSComctlLib.Node
Dim MlngLastSearchIndex As Long
Dim MbolIconeGrandi As Boolean

Public Enum setTipoIconeAlbero
    enmLeggiDaProfilo = 0
    enmIconePiccole = 1
    enmIconeGrandi = 2
End Enum



Private Sub AddGroupItemsTw(objParent As MSXML2.IXMLDOMNode, objTwGroupParent As MSComctlLib.Node)
    Dim objItem As MSXML2.IXMLDOMNode
    Dim objTwGroup As MSComctlLib.Node
    For Each objItem In objParent.childNodes
        If objItem.nodeName = "group" Then   'Sottogruppo
            Set objTwGroup = TwPreferiti.Nodes.Add(objTwGroupParent, tvwChild, , objItem.Attributes.getNamedItem("name").text, "CartCh", "CartAp")
            objTwGroup.Tag = "G"
            Call AddGroupItemsTw(objItem, objTwGroup)
            objTwGroup.Expanded = (objItem.Attributes.getNamedItem("expanded").text = 1)
        Else
            TwPreferiti.Nodes.Add objTwGroupParent.Index, tvwChild, objItem.Attributes.getNamedItem("key").text, objItem.text, objItem.Attributes.getNamedItem("image").text, objItem.Attributes.getNamedItem("image").text
        End If
    Next

End Sub

Private Sub AddGroupItems(oDocument As MSXML2.DOMDocument, objGroupNode As MSXML2.IXMLDOMNode, objParent As MSComctlLib.Node)
    If objParent.children > 0 Then
        Dim objNode As MSComctlLib.Node
        Dim objItem As MSXML2.IXMLDOMNode
        Dim objAttr As MSXML2.IXMLDOMNode
        Dim objGroup As MSXML2.IXMLDOMNode
        Set objNode = objParent.Child
        Do While Not (objNode Is Nothing)
            If objNode.Tag = "G" Then   'Gruppo
                Set objGroup = oDocument.createNode(MSXML2.NODE_ELEMENT, "group", "")
                Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "name", "")
                objAttr.nodeValue = objNode.text
                objGroup.Attributes.setNamedItem objAttr
                Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "expanded", "")
                If objNode.Expanded Then
                    objAttr.nodeValue = 1
                Else
                    objAttr.nodeValue = 0
                End If
                objGroup.Attributes.setNamedItem objAttr
                Call AddGroupItems(oDocument, objGroup, objNode)
                objGroupNode.appendChild objGroup
            Else
                Set objItem = oDocument.createNode(MSXML2.NODE_ELEMENT, "item", "")
                'Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "name", "")
                'objAttr.nodeValue = objNode.text
                'objItem.Attributes.setNamedItem objAttr
                objItem.text = objNode.text
                Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "key", "")
                objAttr.nodeValue = objNode.key
                objItem.Attributes.setNamedItem objAttr
                Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "image", "")
                objAttr.nodeValue = objNode.Image
                objItem.Attributes.setNamedItem objAttr
                objGroupNode.appendChild objItem
            End If
            Set objNode = objNode.Next
        Loop
    End If
End Sub

Public Sub CambiaRisorse()
    CommandBars(2).FindControl(, ID_TLBSEARCH_TEXT).Caption = MXNU.CaricaStringaRes(10346)
    CommandBars(2).FindControl(, ID_TLBSEARCH_TEXT).ToolTipText = MXNU.CaricaStringaRes(10346)
    CommandBars(2).FindControl(, ID_TLBSEARCH_FILTER).ToolTipText = MXNU.CaricaStringaRes(11739)
    CommandBars(2).FindControl(, ID_TLBSEARCH_NODESEARCH).ToolTipText = MXNU.CaricaStringaRes(11740)
    ShortcutCaption1.Caption = MXNU.CaricaStringaRes(11995)
    ShortcutCaption2.Caption = MXNU.CaricaStringaRes(11996)
End Sub

Public Sub CaricaPreferitiUtente()
    Dim oDocument As MSXML2.DOMDocument
    Dim objItem As MSXML2.IXMLDOMNode
    Dim objGroup As MSXML2.IXMLDOMNode
    Dim objRoot As MSXML2.IXMLDOMNode
    'Dim objNodeList As MSXML2.IXMLDOMNodeList
    Dim objNode As MSComctlLib.Node
    Dim objTwGroup As MSComctlLib.Node
    
    TwPreferiti.Nodes.Clear
    If Dir$(MXNU.PercorsoPreferenze & "\Preferiti_" & MXNU.UtenteAttivo & ".xml", vbNormal) <> "" Then
        Set oDocument = New MSXML2.DOMDocument
        If oDocument.Load(MXNU.PercorsoPreferenze & "\Preferiti_" & MXNU.UtenteAttivo & ".xml") Then
            Set objRoot = oDocument.documentElement   'oDocument.selectSingleNode("menu")
            'Set objNodeList = objRoot.childNodes
            For Each objGroup In objRoot.childNodes
                If objGroup.nodeName = "group" Then
                    Set objTwGroup = TwPreferiti.Nodes.Add(, tvwNext, , objGroup.Attributes.getNamedItem("name").text, "CartCh", "CartAp")
                    objTwGroup.Tag = "G"
                    Call AddGroupItemsTw(objGroup, objTwGroup)
                    objTwGroup.Expanded = (objGroup.Attributes.getNamedItem("expanded").text = 1)
                Else
                    TwPreferiti.Nodes.Add , tvwNext, objGroup.Attributes.getNamedItem("key").text, objGroup.text, objGroup.Attributes.getNamedItem("image").text, objGroup.Attributes.getNamedItem("image").text
                End If
            Next
        End If
        Set objItem = Nothing
        Set objGroup = Nothing
        Set objRoot = Nothing
        Set objNode = Nothing
        'Set objNodeList = Nothing
        Set objTwGroup = Nothing
        Set oDocument = Nothing
    End If
    
End Sub

Private Sub CreateToolBar()
    Dim tlbPrinc As XtremeCommandBars.CommandBar
    Dim Btn As XtremeCommandBars.CommandBarControl
    Dim BtnCmb As XtremeCommandBars.CommandBarComboBox
    
    
    CommandBars.GlobalSettings.App = App
    CommandBars.EnableCustomization False
    CommandBars.ActiveMenuBar.Visible = False
        
    'ToolBar Principale
    Set tlbPrinc = CommandBars.Add("SearchBar", xtpBarTop)
    With tlbPrinc.Controls
        Set BtnCmb = .Add(xtpControlComboBox, ID_TLBSEARCH_TEXT, MXNU.CaricaStringaRes(10346))    'Trova
        BtnCmb.IconId = 1000000
        BtnCmb.width = 170
        BtnCmb.Style = xtpComboLabel
        BtnCmb.DropDownListStyle = True
                
        'Set Btn = .Add(xtpControlLabel, 0, Space(1))
        Set Btn = .Add(xtpControlButton, ID_TLBSEARCH_FILTER, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImgLstModuli, "SearchFilter")
        Btn.Checked = True
        Btn.BeginGroup = True
        Btn.ToolTipText = MXNU.CaricaStringaRes(11739)
        Set Btn = .Add(xtpControlButton, ID_TLBSEARCH_NODESEARCH, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImgLstModuli, "SearchNode")
        Btn.ToolTipText = MXNU.CaricaStringaRes(11740)
    End With
    tlbPrinc.ContextMenuPresent = False
    tlbPrinc.ModifyStyle XTP_CBRS_GRIPPER, 0
    tlbPrinc.EnableDocking xtpFlagAlignTop
    CommandBars.Options.ShowExpandButtonAlways = False
    
    ShortcutCaption1.Caption = MXNU.CaricaStringaRes(11995)
    ShortcutCaption2.Caption = MXNU.CaricaStringaRes(11996)
End Sub

Private Function EsisteElementoAlbero(colPar As MSComctlLib.Nodes, vntElemento As Variant) As Boolean
    Dim bolDum As Boolean
    
    On Local Error GoTo EsisteElementoAlberoErr
    EsisteElementoAlbero = True
    bolDum = IsObject(colPar(vntElemento))
    
FineEsisteElementoAlbero:
    On Local Error GoTo 0
    Exit Function
    
EsisteElementoAlberoErr:
    EsisteElementoAlbero = False
    Resume FineEsisteElementoAlbero

End Function

Public Property Get IconeGrandi() As Boolean
    IconeGrandi = MbolIconeGrandi
End Property

Public Sub ImpostaImageList(Optional ByVal TipoIcone As setTipoIconeAlbero = enmLeggiDaProfilo)
    Dim colIconeModuli As Collection
    Dim colIconeSearch As Collection
    Dim colIconePreferiti As Collection
    Dim i As Long
    If TipoIcone = enmLeggiDaProfilo Then
        Dim f As Integer, strRiga As String
        'TipoIcone = Cast(MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", MXNU.UtenteSistema, "TipoIconeModuli", enmIconePiccole), vbLong)
        If Dir(MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini", vbNormal) <> "" Then
            f = FreeFile
            Open MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini" For Input As #f
            Do While Not EOF(f)
                Line Input #f, strRiga
                If strRiga <> "" Then
                    If InStr(LCase(strRiga), "iconegrandi=") <> 0 Then
                        If Val(Mid(strRiga, InStr(strRiga, "=") + 1)) = 0 Then
                            TipoIcone = enmIconePiccole
                        Else
                            TipoIcone = enmIconeGrandi
                        End If
                        Exit Do
                    End If
                End If
            Loop
            Close #f
        End If
        If TipoIcone = enmLeggiDaProfilo Then TipoIcone = enmIconePiccole
    End If
    If TrwModuli.Nodes.Count > 0 Then
        GoSub LeggiIconeNodi
    End If
    If TipoIcone = enmIconeGrandi Then
        Set TrwModuli.ImageList = ImgLstModuli
        Set TwPreferiti.ImageList = ImgLstModuli
        Set TrVSearch.ImageList = ImgLstModuli
    ElseIf TipoIcone = enmIconePiccole Then
        Set TrwModuli.ImageList = ImgLstModuli16x16
        Set TwPreferiti.ImageList = ImgLstModuli16x16
        Set TrVSearch.ImageList = ImgLstModuli16x16
    End If
    If TrwModuli.Nodes.Count > 0 Then
        GoSub ImpostaIconeNodi
    End If
    MbolIconeGrandi = (TipoIcone = enmIconeGrandi)
    If Not MbolIconeGrandi Then
        'Altrimenti passando da Icone Grandi a Icone Piccole resta l'indentazione delle icone a 24x24
        TrwModuli.Indentation = 200.126
        TwPreferiti.Indentation = 200.126
        TrVSearch.Indentation = 200.126
    End If
    
esci_ImpostaImageList:
    Set colIconeModuli = Nothing
    Set colIconeSearch = Nothing
    Set colIconePreferiti = Nothing
    Exit Sub
    
LeggiIconeNodi:
    Set colIconeModuli = New Collection
    Set colIconePreferiti = New Collection
    Set colIconeSearch = New Collection
    For i = 1 To TrwModuli.Nodes.Count
        colIconeModuli.Add TrwModuli.Nodes(i).Image
    Next
    For i = 1 To TrVSearch.Nodes.Count
        colIconeSearch.Add TrVSearch.Nodes(i).Image
    Next
    For i = 1 To TwPreferiti.Nodes.Count
        colIconePreferiti.Add TwPreferiti.Nodes(i).Image
    Next
Return

ImpostaIconeNodi:
    For i = 1 To TrwModuli.Nodes.Count
        TrwModuli.Nodes(i).Image = colIconeModuli(i)
    Next
    For i = 1 To TrVSearch.Nodes.Count
        TrVSearch.Nodes(i).Image = colIconeSearch(i)
    Next
    For i = 1 To TwPreferiti.Nodes.Count
        TwPreferiti.Nodes(i).Image = colIconePreferiti(i)
    Next
Return
    
End Sub

Private Sub Preferiti_AddItem()
    Dim objNewNode As MSComctlLib.Node
    
    'Rif. A#10636
    If Not EsisteElementoAlbero(TwPreferiti.Nodes, TrwModuli.SelectedItem.key) Then
        Set objNewNode = TwPreferiti.Nodes.Add(, tvwLast, TrwModuli.SelectedItem.key, TrwModuli.SelectedItem.text, TrwModuli.SelectedItem.Image, TrwModuli.SelectedItem.SelectedImage)
        objNewNode.EnsureVisible
    Else
        MXNU.MsgBoxEX 3199, vbInformation, 1007, Array(TrwModuli.SelectedItem.text)
    End If
    Set objNewNode = Nothing
    
    'Rif. A#10635 - Salvo i preferiti ad ogni modifica
    Call SalvaPreferitiUtente
End Sub

Public Sub SalvaPreferitiUtente()
    Dim oDocument As MSXML2.DOMDocument
    Dim oPi As MSXML2.IXMLDOMProcessingInstruction
    Dim objRoot As MSXML2.IXMLDOMNode
    Dim objGroup As MSXML2.IXMLDOMNode
    Dim objItem As MSXML2.IXMLDOMNode
    Dim objAttr As MSXML2.IXMLDOMNode
    Dim objNode As MSComctlLib.Node
    
    'RIF. AN. #11607
    If hndDBArchivi.ConnessioneR.State = 0 Then Exit Sub
    
    If TwPreferiti.Nodes.Count = 0 Then
        On Local Error Resume Next
        Kill MXNU.PercorsoPreferenze & "\Preferiti_" & MXNU.UtenteAttivo & ".xml"
        On Local Error GoTo 0
        Exit Sub
    End If
    On Local Error GoTo err_SalvaPreferiti
    Set oDocument = New MSXML2.DOMDocument
    
    'processing istruction xml
    Set oPi = oDocument.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    Call oDocument.appendChild(oPi)
    
    
    Set objRoot = oDocument.createNode(MSXML2.NODE_ELEMENT, "menu", "")
    Call oDocument.appendChild(objRoot)
    
    Set objNode = TwPreferiti.Nodes(1)
    Do While Not (objNode Is Nothing)
        If objNode.Tag = "G" Then   'Gruppo
            Set objGroup = oDocument.createNode(MSXML2.NODE_ELEMENT, "group", "")
            Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "name", "")
            objAttr.nodeValue = objNode.text
            objGroup.Attributes.setNamedItem objAttr
            Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "expanded", "")
            If objNode.Expanded Then
                objAttr.nodeValue = 1
            Else
                objAttr.nodeValue = 0
            End If
            objGroup.Attributes.setNamedItem objAttr
            Call AddGroupItems(oDocument, objGroup, objNode)
            objRoot.appendChild objGroup
        Else
            Set objItem = oDocument.createNode(MSXML2.NODE_ELEMENT, "item", "")
            'Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "name", "")
            'objAttr.nodeValue = objNode.text
            'objItem.Attributes.setNamedItem objAttr
            objItem.text = objNode.text
            Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "key", "")
            objAttr.text = objNode.key
            objItem.Attributes.setNamedItem objAttr
            Set objAttr = oDocument.createNode(MSXML2.NODE_ATTRIBUTE, "image", "")
            objAttr.nodeValue = objNode.Image
            objItem.Attributes.setNamedItem objAttr
            objRoot.appendChild objItem
        End If
        Set objNode = objNode.Next
    Loop
    oDocument.Save MXNU.PercorsoPreferenze & "\Preferiti_" & MXNU.UtenteAttivo & ".xml"
    
esci_SalvaPreferiti:
    On Local Error GoTo 0
    Set objItem = Nothing
    Set objGroup = Nothing
    Set objAttr = Nothing
    Set objRoot = Nothing
    Set oPi = Nothing
    Set objNode = Nothing
    Set oDocument = Nothing
    Exit Sub

err_SalvaPreferiti:
    MXNU.MsgBoxEX 1009, vbCritical, 1007, Array("SalvaPreferitiUtente", Err.Number, Err.Description)
    Resume esci_SalvaPreferiti
    
    Resume
    
End Sub





Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim GroupNode As MSComctlLib.Node
    Select Case Control.ID
        Case ID_POPUP_RUNPROG
            'Call TrwModuli_DblClick
            If Me.ActiveControl.NAME = "TrwModuli" Then
                Call TrwModuli_DblClick
            Else
                Call TwPreferiti_DblClick
            End If
        Case ID_POPUP_ADDPREFERITI
            Call Preferiti_AddItem
        Case ID_POPUP_ADDGROUP
            Call Preferiti_AddGroup
        Case ID_POPUP_RENAME
            SendKeys "{F2}"   'Non funziona il rename chiamando direttamente la funzione Preferiti_Rename (?!?)
            'Call Preferiti_Rename
        Case ID_POPUP_DELGROUP
            Call Preferiti_DelGroup
        Case ID_POPUP_DELITEM
            Call Preferiti_DelItem
        Case ID_TLBSEARCH_TEXT
            Call DoSearch(CommandBars(2).FindControl(, ID_TLBSEARCH_FILTER).Checked)
            Dim BtnCmb As XtremeCommandBars.CommandBarComboBox
            Dim i As Long
            Dim strSearch As String, bolTrovato As Boolean
            
            Set BtnCmb = Control
            strSearch = BtnCmb.text
            If strSearch <> "" Then
                'Aggiungo la chiave di ricerca al combo, per eventuale riutilizzo
                bolTrovato = False
                For i = 1 To BtnCmb.ListCount
                    If StrComp(strSearch, BtnCmb.List(i), vbTextCompare) = 0 Then
                        bolTrovato = True
                        Exit For
                    End If
                Next i
                If Not bolTrovato Then
                    BtnCmb.addItem strSearch
                End If
            End If
            On Local Error Resume Next
            Control.SetFocus
            SendKeys "{ENTER}"   'Per posizionare il focus sulla parte editabile
            On Local Error GoTo 0
        Case ID_TLBSEARCH_FILTER
            CommandBars(2).FindControl(, ID_TLBSEARCH_NODESEARCH).Checked = False
            Control.Checked = True
        Case ID_TLBSEARCH_NODESEARCH
            CommandBars(2).FindControl(, ID_TLBSEARCH_FILTER).Checked = False
            Control.Checked = True
    End Select
                    
End Sub


Private Sub DoSearch(ByVal HideNoMatch As Boolean)
    Dim colNodiAlbero As New Collection
    Dim i As Long
    Dim bolValido As Boolean
    Dim strSearch As String
    Call VisitaAlbero(TrwModuli, colNodiAlbero)
    strSearch = CommandBars(2).FindControl(, ID_TLBSEARCH_TEXT).text
    If Not HideNoMatch Then
        For i = 1 To colNodiAlbero.Count
            With TrwModuli
                If MlngLastSearchIndex = 0 Then
                    bolValido = True
                Else
                    bolValido = (.Nodes(colNodiAlbero(i)).Index > MlngLastSearchIndex)
                End If
                'Debug.Print .Nodes(colNodiAlbero(i)).text
                If bolValido Then
                    If LCase(.Nodes(colNodiAlbero(i)).text) Like LCase(strSearch) Then
                        .Nodes(colNodiAlbero(i)).Selected = True
                        .Nodes(colNodiAlbero(i)).EnsureVisible
                        MlngLastSearchIndex = .Nodes(colNodiAlbero(i)).Index
                        Exit For
                    End If
                End If
            End With
        Next i
        If i > colNodiAlbero.Count Then MlngLastSearchIndex = 0
    Else
        If strSearch = "" Then
            'Se sbianco il campo di ricerca ripristino l'albero originale
            TrVSearch.Visible = False
            TrwModuli.Visible = True
        Else
            Dim objNode As MSComctlLib.Node
            TrVSearch.Nodes.Clear
            For i = 1 To colNodiAlbero.Count
                With TrwModuli
                    If LCase(.Nodes(colNodiAlbero(i)).text) Like LCase(strSearch) Then
                        Set objNode = .Nodes(colNodiAlbero(i))
                        Call AggiungiPadre(objNode)
                        If Not (objNode.Parent Is Nothing) Then
                            TrVSearch.Nodes.Add objNode.Parent.key, tvwChild, objNode.key, objNode.text, objNode.Image, objNode.SelectedImage
                            TrVSearch.Nodes(objNode.key).EnsureVisible
                        End If
                    ElseIf Not (.Nodes(colNodiAlbero(i)).Parent Is Nothing) Then
                        If LCase(.Nodes(colNodiAlbero(i)).Parent.text) Like LCase(strSearch) Then
                            TrVSearch.Nodes.Add .Nodes(colNodiAlbero(i)).Parent.key, tvwChild, .Nodes(colNodiAlbero(i)).key, .Nodes(colNodiAlbero(i)).text, .Nodes(colNodiAlbero(i)).Image, .Nodes(colNodiAlbero(i)).SelectedImage
                            TrVSearch.Nodes(.Nodes(colNodiAlbero(i)).key).EnsureVisible
                        End If
                    End If
                End With
            Next i
            If TrVSearch.Nodes.Count > 0 Then
                TrVSearch.Left = TrwModuli.Left
                TrVSearch.Top = TrwModuli.Top
                TrVSearch.Height = TrwModuli.Height
                TrVSearch.width = TrwModuli.width
                TrVSearch.Visible = True
                TrVSearch.ZOrder 0
                TrwModuli.Visible = False
            Else
                TrwModuli.Visible = True
                TrVSearch.Visible = False
            End If
        End If
    End If
    Set colNodiAlbero = Nothing
End Sub

Private Sub AggiungiPadre(objNode As MSComctlLib.Node)
    If Not (objNode.Parent Is Nothing) Then
        Dim objPadre As MSComctlLib.Node
        Set objPadre = objNode.Parent
        On Local Error Resume Next
        If Not (objPadre.Parent Is Nothing) Then
            Call AggiungiPadre(objPadre)
            TrVSearch.Nodes.Add objPadre.Parent.key, tvwChild, objNode.Parent.key, objNode.Parent.text, objNode.Parent.Image, objNode.Parent.SelectedImage
        Else
            TrVSearch.Nodes.Add , , objNode.Parent.key, objNode.Parent.text, objNode.Parent.Image, objNode.Parent.SelectedImage
        End If
        On Local Error GoTo 0
    End If
End Sub

Private Sub VisitaAlbero(Albero As MSComctlLib.TreeView, colNodiAlbero As Collection)
    Dim i&
    For i = 1 To colNodiAlbero.Count
        colNodiAlbero.Remove 1
    Next i
    colNodiAlbero.Add Albero.Nodes(1).key, Albero.Nodes(1).key
    Call VisitaNodo(Albero.Nodes(1), colNodiAlbero)
End Sub

Private Sub VisitaNodo(ByVal objParent As Node, colNodiAlbero As Collection)
Dim objNode As Node

    If objParent.children > 0 Then
        Set objNode = objParent.Child
        Do While Not (objNode Is Nothing)
            'Debug.Print objNode.Text
            If Not EsisteElementoCollection(colNodiAlbero, objNode.key) Then
                colNodiAlbero.Add objNode.key, objNode.key
                VisitaNodo objNode, colNodiAlbero
                Set objNode = objNode.Next
                If objNode Is Nothing Then
                    Set objNode = objParent.Next
                    If Not (objNode Is Nothing) Then Set objParent = objNode
                End If
            Else
                Set objNode = Nothing
            End If
        Loop
    End If
End Sub

Private Sub CommandBars_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Top = ShortcutCaption1.Height
End Sub


Private Sub Form_Load()
    'MWSplitter1.BackColor = vbActiveTitleBar
    Call CambiaCharSet(Me)
    MWSplitter1.BackColor = CommandBars.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
    Call ImpostaImageList
    Call LoadIniMenu(TrwModuli)
    '[11/04/2011] Rimozione Chiave Hardware
'    'richiesta di aggiornamento del file licenze al server TSE
'    'RIF.TECEUROLAB (09/09/2010)
'    If (MXNU.ControlloModulichiave(modProcessWatcher) <> 0) Then
'        Call MXNU.AggiornaModuliTSE
'    End If
    
    Call CaricaPreferitiUtente
    Call CreateToolBar
    
End Sub


Private Sub Form_Resize()
    upPanel.width = Me.width
    bottomPanel.width = Me.width
    MWSplitter1.width = Me.width
    If Me.Height < MWSplitter1.Top Then    '??? succede in caricamento della form
        bottomPanel.Height = Me.Height
    Else
        On Local Error Resume Next
        bottomPanel.Height = Me.Height - (MWSplitter1.Top + MWSplitter1.Height) + 15
        On Local Error GoTo 0
    End If
    CommandBars.Item(2).FindControl(, ID_TLBSEARCH_TEXT).width = (Me.width / Screen.TwipsPerPixelX) - 60
    CommandBars.RecalcLayout
End Sub

Public Property Let ModuloAttivo(NuovoModulo As Variant)
    If NuovoModulo = "*DA_INI*" Then
        On Local Error Resume Next
        Dim strMenu$
        With TrwModuli
            strMenu = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\Mw.ini", MXNU.UtenteSistema, "ModuloAttivo", "")
            If strMenu <> "" Then
                Set .SelectedItem = .Nodes(strMenu)
                TrwModuli_NodeClick .Nodes(strMenu)
                .Nodes(strMenu).EnsureVisible
                strMenu = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\Mw.ini", MXNU.UtenteSistema, "MenuAttivo", "")  'Sostituito UtenteAttivo con UtenteSistema per Anomalia 10625
                If strMenu <> "" Then
                    Set .SelectedItem = .Nodes(strMenu)
                    TrwModuli_DblClick
                End If
            Else
                TrwModuli_NodeClick TrwModuli.Nodes("Metodo98")
                TrwModuli.Nodes("Metodo98").EnsureVisible
            End If
        End With
        On Local Error GoTo 0
    Else
        TrwModuli_NodeClick TrwModuli.Nodes(NuovoModulo)
        TrwModuli.Nodes(NuovoModulo).EnsureVisible
    End If
    
    If Not metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden Then
        If frmModuli2005.ShortcutBar1.Selected.ID = ID_BAR_PROGMODULES Then
            Call DaiFocusAlberoModuli
        End If
    End If
    
End Property

Private Sub Form_Unload(Cancel As Integer)
    Call SalvaPreferitiUtente
    Call ScaricaMenu
    Set frmMenu = Nothing
End Sub

'Private Sub twMenu_NodeClick(ByVal Node As MSComctlLib.Node)
'    Dim strMenu As String
'    Static strModuloAttuale As String
'    On Local Error Resume Next
'    If Node.key = "MetodoXP" Then Exit Sub
'    DoEvents
'    If UCase(McolMenu(Node.key).Modulo) <> UCase(strModuloAttuale) Then
'        Call CaricaMenu(McolMenu(Node.key).Modulo)
'        strModuloAttuale = McolMenu(Node.key).Modulo
'        If MDIfrmMenu.DockingManager.Panes(3).Hidden Then MDIfrmMenu.DockingManager.Panes(3).Selected = True
'    End If
'
'End Sub


Private Sub MWSplitter1_AfterResizeObject(objResized As Object)
    'Controllo che l'utente non draggi oltre la linea di inizio dei pulsanti della Navigation Bar altrimenti poi non riesce più a ridimensionare i preferiti
    'Set ModShortcutBar = frmModuli2005.ShortcutBar1
    If objResized Is bottomPanel Then
        RecalcPreferiti
    End If
End Sub

Public Sub RecalcPreferiti()
Dim lngLimiteTop As Long
        Dim r As RECT
        
        On Local Error Resume Next
        Static bolResizing As Boolean

        If bolResizing Then Exit Sub
        
        Call GetClientRect(frmModuli2005.ShortcutBar1.hwnd, r)
        
        lngLimiteTop = (r.Top * Screen.TwipsPerPixelY) + (r.Bottom * Screen.TwipsPerPixelY) - (frmModuli2005.ShortcutBar1.ExpandedLinesCount * 31 * Screen.TwipsPerPixelY) - (7 * Screen.TwipsPerPixelY)
        
        If (MWSplitter1.Top + MWSplitter1.Height + ShortcutCaption2.Height) >= lngLimiteTop Then
            bolResizing = True
            MWSplitter1.SplitterTop = ((lngLimiteTop + MWSplitter1.Height + ShortcutCaption2.Height) \ Screen.TwipsPerPixelY)
            bolResizing = False
        End If
        On Local Error GoTo 0
End Sub


Private Sub TrVSearch_Collapse(ByVal Node As MSComctlLib.Node)
    Call TrwModuli_Collapse(Node)
End Sub


Private Sub TrVSearch_DblClick()
    TrwModuli.SelectedItem = TrVSearch.SelectedItem
    Call TrwModuli_DblClick
End Sub


Private Sub TrVSearch_Expand(ByVal Node As MSComctlLib.Node)
    Call TrwModuli_Expand(Node)
End Sub


Private Sub TrVSearch_KeyPress(KeyAscii As Integer)
    TrwModuli.SelectedItem = TrVSearch.SelectedItem
    Call TrwModuli_KeyPress(KeyAscii)
End Sub


Private Sub TrVSearch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim NodoPuntato As Node
    If Shift = vbCtrlMask Then
        Set NodoPuntato = TrVSearch.HitTest(x, y)
        If Not (NodoPuntato Is Nothing) Then
            If NodoPuntato.children > 0 Then Exit Sub   'Nei preferiti si mettono solo nodi foglia
            If Button = vbLeftButton Then
                TrVSearch.Drag vbBeginDrag
                Set NodoDraggato = NodoPuntato
            End If
            Set NodoPuntato = Nothing
        End If
    End If

End Sub


Private Sub TrVSearch_NodeClick(ByVal Node As MSComctlLib.Node)
    Call TrwModuli_NodeClick(Node)
End Sub


Private Sub TrwModuli_Collapse(ByVal Node As MSComctlLib.Node)
    If Node.Image = "CartAp" Then
        Node.Image = "CartCh"
    End If
End Sub

Private Sub TrwModuli_DblClick()
    If (IsNodoFoglia(TrwModuli.SelectedItem, False)) Then
        AttivaVoceMenu TrwModuli.SelectedItem.key
    End If
End Sub

Private Function IsNodoFoglia(nodX As Node, bolPreferiti As Boolean) As Boolean
    Dim twMenu As MSComctlLib.TreeView
    If bolPreferiti Then
        Set twMenu = TwPreferiti
    Else
        Set twMenu = TrwModuli
    End If
    'IsNodoFoglia = (TrwModuli.SelectedItem.Child Is Nothing)
    IsNodoFoglia = (twMenu.SelectedItem.Child Is Nothing)
End Function

Private Sub TrwModuli_Expand(ByVal Node As MSComctlLib.Node)
    If Node.Image = "CartCh" Then
        Node.Image = "CartAp"
    End If
End Sub

Private Sub TrwModuli_KeyPress(KeyAscii As Integer)
    'gestione del tasto invio sul nodo foglia
    If (KeyAscii = vbKeyReturn) Then
        If (IsNodoFoglia(TrwModuli.SelectedItem, False)) Then
            Call TrwModuli_DblClick
        End If
    End If
End Sub


Private Sub TrwModuli_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Dim NodoPuntato As MSComctlLib.Node
        Set NodoPuntato = TrwModuli.HitTest(x, y)
        If Not (NodoPuntato Is Nothing) Then
            If IsNodoFoglia(NodoPuntato, False) Then
                TrwModuli.SelectedItem = NodoPuntato
                CommandBars.AddImageList ImgLstModuli16x16
                Dim PopupBar As CommandBar
                Dim Btn As CommandBarControl
                
                Set PopupBar = CommandBars.Add("Popup", xtpBarPopup)
                With PopupBar.Controls
                    Set Btn = .Add(xtpControlButton, ID_POPUP_RUNPROG, "&Esegui...")
                    Btn.IconId = ImgListKey2ImgListIdx(ImgLstModuli16x16, "RunProg")
                    Set Btn = .Add(xtpControlButton, ID_POPUP_ADDPREFERITI, "&Aggiungi ai Preferiti")
                    Btn.IconId = ImgListKey2ImgListIdx(ImgLstModuli16x16, "Preferiti")
                End With
                PopupBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub TrwModuli_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim NodoPuntato As Node
    If Shift = vbCtrlMask Then
        Set NodoPuntato = TrwModuli.HitTest(x, y)
        If Not (NodoPuntato Is Nothing) Then
            If NodoPuntato.children > 0 Then Exit Sub   'Nei preferiti si mettono solo nodi foglia
            If Button = vbLeftButton Then
                TrwModuli.Drag vbBeginDrag
                Set NodoDraggato = NodoPuntato
            End If
            Set NodoPuntato = Nothing
        End If
    End If
End Sub

Private Sub TrwModuli_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strMenu As String
    Static strModuloAttuale As String
    On Local Error Resume Next
    'If Node.key = "Metodo2005" Then Exit Sub
    DoEvents
    If UCase(McolMenu(Node.key).Modulo) <> UCase(strModuloAttuale) Or Node.key = "Metodo98" Then  'ricarico comunque il modulo principale (Metodo98) in quanto se si cambia la lingua attiva ed il mod. selezionato era comuni, restano le voci di menu con la lingua precedente
        Dim ctlLblModulo As CommandBarControl
        Call CaricaMenuXCB(McolMenu(Node.key).Modulo)
        strModuloAttuale = McolMenu(Node.key).Modulo
        'If metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden Then metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Selected = True
        'On Local Error Resume Next
        Set ctlLblModulo = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_LBLMODULO)
        If strModuloAttuale <> "Comuni" Then
            If strModuloAttuale <> "Metodo98" Then
                Err.Clear
                ctlLblModulo.Caption = TrwModuli.Nodes(strModuloAttuale).text
                If Err.Number <> 0 Then
                    ctlLblModulo.Caption = strModuloAttuale
                End If
            Else
                ctlLblModulo.Caption = McolMenu(Node.key).Caption
            End If
            'If ImgListKey2ImgListIdx(ImgLstModuli, strModuloAttuale) > 0 Then
            '    ctlLblModulo.IconId = ImgListKey2ImgListIdx(ImgLstModuli, strModuloAttuale)
            'Else
                'ctlLblModulo.IconId = ImgListKey2ImgListIdx(ImgLstModuli, TrwModuli.SelectedItem.Image)
                'ctlLblModulo.IconId = ImgListKey2ImgListIdx(ImgLstModuli, strModuloAttuale)
                ctlLblModulo.IconId = ImgListKey2ImgListIdx(ImgLstModuli, TrwModuli.Nodes(strModuloAttuale).Image)
            'End If
        Else
            ctlLblModulo.Caption = TrwModuli.Nodes(McolMenu(Node.key).Modulo).text
            ctlLblModulo.IconId = ImgListKey2ImgListIdx(ImgLstModuli, "Metodo98")
        End If
        metodo.FindCommandBar(ID_TLB_QUALITY).Visible = (LCase(strModuloAttuale) = "quality")
        'metodo.CommandBars.RecalcLayout
        If metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden Then metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Selected = True
        'On Local Error GoTo 0
        
        'Anomalia 11668
        If LCase(strModuloAttuale) <> "anabilancio" And LCase(strModuloAttuale) <> "tesoreria" Then
            Call AggiornaStatusBar
        End If
    End If

End Sub


Private Sub TwPreferiti_Collapse(ByVal Node As MSComctlLib.Node)
    If Node.Image = "CartAp" Then
        Node.Image = "CartCh"
        Node.SelectedImage = "CartCh"
    End If
End Sub

Private Sub TwPreferiti_DblClick()
    If Not (TwPreferiti.SelectedItem Is Nothing) Then   'Anomalia 8506
        If (IsNodoFoglia(TwPreferiti.SelectedItem, True)) Then
            If TwPreferiti.SelectedItem.Tag <> "G" Then
                AttivaVoceMenu TwPreferiti.SelectedItem.key
            End If
        End If
    End If
End Sub

Private Sub TwPreferiti_DragDrop(Source As Control, x As Single, y As Single)
    Dim objInsNode As MSComctlLib.Node
    Dim objNewNode As MSComctlLib.Node
    If TwPreferiti.Nodes.Count > 0 Then
        Set objInsNode = TwPreferiti.HitTest(x, y)
        'If objInsNode Is Nothing Then Set objInsNode = TwPreferiti.Nodes(TwPreferiti.Nodes.Count)
        On Local Error GoTo err_AddPreferiti
        If Not (objInsNode Is Nothing) Then
            If objInsNode.key = NodoDraggato.key Then Exit Sub
            objInsNode.Selected = True
            If objInsNode.Tag = "G" Then  'è un gruppo
                Set objNewNode = TwPreferiti.Nodes.Add(objInsNode.Index, tvwChild, NodoDraggato.key, NodoDraggato.text, NodoDraggato.Image, NodoDraggato.SelectedImage)
            ElseIf Not (objInsNode.Parent Is Nothing) Then
                Set objNewNode = TwPreferiti.Nodes.Add(objInsNode.Parent.Index, tvwChild, NodoDraggato.key, NodoDraggato.text, NodoDraggato.Image, NodoDraggato.SelectedImage)
            Else
                Set objNewNode = TwPreferiti.Nodes.Add(objInsNode.Index, tvwNext, NodoDraggato.key, NodoDraggato.text, NodoDraggato.Image, NodoDraggato.SelectedImage)
            End If
            'TwPreferiti.Nodes.Add , , NodoDraggato.key, NodoDraggato.Text, NodoDraggato.Image, NodoDraggato.SelectedImage
            'Set objInsNode = TwPreferiti.Nodes.Add("Root", tvwChild, NodoDraggato.key, NodoDraggato.Text, NodoDraggato.Image, NodoDraggato.SelectedImage)
            objNewNode.EnsureVisible
        Else
            Set objNewNode = TwPreferiti.Nodes.Add(, tvwNext, NodoDraggato.key, NodoDraggato.text, NodoDraggato.Image, NodoDraggato.SelectedImage)
            objNewNode.EnsureVisible
        End If
    Else
        TwPreferiti.Nodes.Add , tvwFirst, NodoDraggato.key, NodoDraggato.text, NodoDraggato.Image, NodoDraggato.SelectedImage
    End If
    
    'Rif. A#10635 - Salvo i preferiti ad ogni modifica
    Call SalvaPreferitiUtente
    
esci_AddPreferiti:
    Exit Sub
    
err_AddPreferiti:
    If Err.Number = 35602 Then    'Chiave non univoca nell'insieme
        MXNU.MsgBoxEX "L'elemento '" & NodoDraggato.text & "' esiste già nei preferiti!", vbExclamation, 1007
    Else
        MXNU.MsgBoxEX "Errore " & Err.Number & ": " & Err.Description, vbExclamation, 1007
    End If
    Resume esci_AddPreferiti
    
End Sub

Private Sub TwPreferiti_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If (Source Is TrwModuli) Or (Source Is TrVSearch) Then
        If TwPreferiti.Nodes.Count > 0 Then
            Dim objInsNode As MSComctlLib.Node
            Set objInsNode = TwPreferiti.HitTest(x, y)
            If Not objInsNode Is Nothing Then
                On Local Error Resume Next
                TwPreferiti.SetFocus
                On Local Error GoTo 0
                Set TwPreferiti.SelectedItem = objInsNode
            End If
        End If
        Source.DragIcon = ImgDrag(1).Picture
    End If

End Sub


Private Sub TwPreferiti_Expand(ByVal Node As MSComctlLib.Node)
    If Node.Image = "CartCh" Then
        Node.Image = "CartAp"
        Node.SelectedImage = "CartAp"
    End If
End Sub

Private Sub TwPreferiti_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (TwPreferiti.SelectedItem Is Nothing) Then
        Select Case KeyCode
            Case vbKeyF2
                If KeyCode = vbKeyF2 Then
                    Call Preferiti_Rename
                End If
            Case vbKeyDelete
                If TwPreferiti.SelectedItem.Tag = "G" Then
                    Call Preferiti_DelGroup
                Else
                    Call Preferiti_DelItem
                End If
'        If KeyCode = vbKeyF2 And TwPreferiti.SelectedItem.Tag = "G" Then
'            TwPreferiti.StartLabelEdit
'        ElseIf KeyCode = vbKeyCancel Then
'            If TwPreferiti.SelectedItem.Tag = "G" Then
'            End If
'
'        End If
        End Select
    End If

End Sub

Private Sub Preferiti_AddGroup()
    Dim strNewGroup As String
    Dim objNodeGrp As MSComctlLib.Node
    strNewGroup = InputBox("Nome nuovo gruppo", "Aggiunta di un nuovo gruppo di preferiti")
    If strNewGroup <> "" Then
        If Not (TwPreferiti.SelectedItem Is Nothing) Then
            If TwPreferiti.SelectedItem.children > 0 Then
                'Set objNodeGrp = TwPreferiti.Nodes.Add(TwPreferiti.SelectedItem, tvwNext, , strNewGroup, "menunew", "menunew")
                Set objNodeGrp = TwPreferiti.Nodes.Add(TwPreferiti.SelectedItem, tvwChild, , strNewGroup, "CartCh", "CartAp")
            ElseIf Not (TwPreferiti.SelectedItem.Parent Is Nothing) Then
                'Set objNodeGrp = TwPreferiti.Nodes.Add(TwPreferiti.SelectedItem.Parent, tvwNext, , strNewGroup, "menunew", "menunew")
                Set objNodeGrp = TwPreferiti.Nodes.Add(TwPreferiti.SelectedItem.Parent, tvwChild, , strNewGroup, "CartCh", "CartAp")
            Else
                Set objNodeGrp = TwPreferiti.Nodes.Add(, tvwNext, , strNewGroup, "CartCh", "CartAp")
            End If
        Else
            Set objNodeGrp = TwPreferiti.Nodes.Add(, tvwNext, , strNewGroup, "CartCh", "CartAp")
        End If
        objNodeGrp.Tag = "G"
    End If
    
    'Rif. A#10635 - Salvo i preferiti ad ogni modifica
    Call SalvaPreferitiUtente
End Sub

Private Sub Preferiti_Rename()
    Dim Node As MSComctlLib.Node
    If Not (TwPreferiti.SelectedItem Is Nothing) Then
        'If TwPreferiti.SelectedItem.children > 0 Then
        'If TwPreferiti.SelectedItem.Tag = "G" Then
            Set Node = TwPreferiti.SelectedItem
        'Else
        '    Set GroupNode = TwPreferiti.SelectedItem.Parent
        'End If
        'If Not (GroupNode Is Nothing) Then
            Node.Selected = True
            Node.EnsureVisible
            DoEvents
            TwPreferiti.StartLabelEdit
            
        'End If
    End If

    'Rif. A#10635 - Salvo i preferiti ad ogni modifica
    Call SalvaPreferitiUtente
End Sub

Private Sub Preferiti_DelGroup()
    If Not (TwPreferiti.SelectedItem Is Nothing) Then
        Dim GroupNode As MSComctlLib.Node
        'If TwPreferiti.SelectedItem.children > 0 Then
        If TwPreferiti.SelectedItem.Tag = "G" Then
            Set GroupNode = TwPreferiti.SelectedItem
        Else
            Set GroupNode = TwPreferiti.SelectedItem.Parent
        End If
        If Not (GroupNode Is Nothing) Then
            If MXNU.MsgBoxEX("Confermi eliminazione gruppo '" & GroupNode.text & "' ?", vbQuestion + vbYesNo, 1007) = vbYes Then
                TwPreferiti.Nodes.Remove GroupNode.Index
            End If
        End If
    End If
    
    'Rif. A#10635 - Salvo i preferiti ad ogni modifica
    Call SalvaPreferitiUtente
End Sub

Private Sub Preferiti_DelItem()
    If Not (TwPreferiti.SelectedItem Is Nothing) Then
        If MXNU.MsgBoxEX("Rimuovere la voce '" & TwPreferiti.SelectedItem.text & "' dai preferiti?", vbQuestion + vbYesNo, 1007) = vbYes Then
            TwPreferiti.Nodes.Remove TwPreferiti.SelectedItem.Index
        End If
    End If
    
    'Rif. A#10635 - Salvo i preferiti ad ogni modifica
    Call SalvaPreferitiUtente
End Sub

Private Sub TwPreferiti_KeyPress(KeyAscii As Integer)
    'gestione del tasto invio sul nodo foglia
    If (KeyAscii = vbKeyReturn) Then
        If (IsNodoFoglia(TwPreferiti.SelectedItem, True)) Then
            Call TwPreferiti_DblClick
        End If
    End If

End Sub

Private Sub TwPreferiti_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Dim NodoPuntato As MSComctlLib.Node
        Set NodoPuntato = TwPreferiti.HitTest(x, y)
        
        If Not (NodoPuntato Is Nothing) Then TwPreferiti.SelectedItem = NodoPuntato
        'CommandBars.AddImageList ImgLstModuli
        CommandBars.AddImageList ImgLstModuli16x16
        Dim PopupBar As CommandBar
        Dim Btn As CommandBarControl
        
        Set PopupBar = CommandBars.Add("Popup", xtpBarPopup)
        With PopupBar.Controls
            If Not (NodoPuntato Is Nothing) Then
                If IsNodoFoglia(NodoPuntato, True) Then
                    Set Btn = .Add(xtpControlButton, ID_POPUP_RUNPROG, "&Esegui...")
                    Btn.IconId = ImgListKey2ImgListIdx(ImgLstModuli16x16, "RunProg")
                End If
            End If
            Set Btn = .Add(xtpControlButton, ID_POPUP_ADDGROUP, "&Aggiungi Gruppo")
            Btn.IconId = ImgListKey2ImgListIdx(ImgLstModuli16x16, "menunew")
            If Not (NodoPuntato Is Nothing) Then
                Set Btn = .Add(xtpControlButton, ID_POPUP_RENAME, "&Rinomina Gruppo/Voce")
                Btn.IconId = ImgListKey2ImgListIdx(ImgLstModuli16x16, "Rename")
                Btn.Enabled = Not (NodoPuntato Is Nothing)
                'If Not (NodoPuntato Is Nothing) Then
                '    Btn.Enabled = Not (NodoPuntato.Parent Is Nothing)
                'Else
                '    Btn.Enabled = False
                'End If
                Set Btn = .Add(xtpControlButton, ID_POPUP_DELGROUP, "&Elimina Gruppo")
                Btn.IconId = ImgListKey2ImgListIdx(ImgLstModuli16x16, "DelGrp")
                Set Btn = .Add(xtpControlButton, ID_POPUP_DELITEM, "Rimuo&vi dai preferiti")
                Btn.IconId = ImgListKey2ImgListIdx(ImgLstModuli16x16, "Del")
                If Not (NodoPuntato Is Nothing) Then
                    Btn.Enabled = (NodoPuntato.Tag <> "G")
                End If
            End If
        End With
        'CommandBars.KeyBindings.Add 0, VK_F2, ID_POPUP_RENAME   '<<< Remmato per anomalia 9211
        PopupBar.ShowPopup
    End If
End Sub


Private Sub upPanel_Resize()
    TrwModuli.Top = 0
    TrwModuli.Height = upPanel.Height
    TrwModuli.width = upPanel.width
    ShortcutCaption1.width = upPanel.width
    TrVSearch.Top = TrwModuli.Top
    TrVSearch.Left = TrwModuli.Left
    TrVSearch.width = TrwModuli.width
    TrVSearch.Height = TrwModuli.Height
End Sub

Private Sub bottomPanel_Resize()
    ShortcutCaption2.Top = 0
    On Local Error Resume Next
    TwPreferiti.Height = bottomPanel.Height - ShortcutCaption2.Height
    On Local Error GoTo 0
    TwPreferiti.Top = ShortcutCaption2.Height '- 30
    TwPreferiti.width = bottomPanel.width
    ShortcutCaption2.width = bottomPanel.width
End Sub

