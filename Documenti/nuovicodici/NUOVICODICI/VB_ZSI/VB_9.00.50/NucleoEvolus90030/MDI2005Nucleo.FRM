VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CODEJOCK.COMMANDBARS.V15.3.1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CODEJOCK.DOCKINGPANE.V15.3.1.OCX"
Object = "{D9CCABDE-1C7B-4CDF-8A3B-882AADC24EDB}#1.0#0"; "MXCTRL.OCX"
Begin VB.MDIForm metodo 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00616161&
   Caption         =   "Metodo Evolus"
   ClientHeight    =   7530
   ClientLeft      =   2640
   ClientTop       =   4395
   ClientWidth     =   14445
   Icon            =   "MDI2005Nucleo.frx":0000
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin MXCtrl.MWSchedaBox SchedaDDE 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LightColor      =   6908265
      ScaleWidth      =   14445
      ScaleHeight     =   435
      Begin VB.TextBox txtDDE 
         Height          =   285
         Left            =   4680
         LinkTimeout     =   100
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   4485
      End
   End
   Begin MSComctlLib.ImageList ImmaginiDesigner 
      Left            =   3540
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":74F2
            Key             =   ""
            Object.Tag             =   "705"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":7A8C
            Key             =   ""
            Object.Tag             =   "702"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":7B5A
            Key             =   ""
            Object.Tag             =   "703"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":7F28
            Key             =   ""
            Object.Tag             =   "708"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":7FE4
            Key             =   ""
            Object.Tag             =   "706"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":83ED
            Key             =   ""
            Object.Tag             =   "704"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":8987
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":8A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":8DEA
            Key             =   ""
            Object.Tag             =   "707"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":9184
            Key             =   "NoVisible"
            Object.Tag             =   "714"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":94ED
            Key             =   "NoEditable"
            Object.Tag             =   "715"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":9856
            Key             =   "PosAltered"
            Object.Tag             =   "716"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":9BBF
            Key             =   "TabAltered"
            Object.Tag             =   "717"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglQuality 
      Left            =   2820
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":9F28
            Key             =   "New"
            Object.Tag             =   "601"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":A47A
            Key             =   "Open"
            Object.Tag             =   "602"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":A9CC
            Key             =   "Save"
            Object.Tag             =   "603"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":AF1E
            Key             =   "OldDelete"
            Object.Tag             =   "604"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":B470
            Key             =   "Preview"
            Object.Tag             =   "605"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":B9C2
            Key             =   "Print"
            Object.Tag             =   "606"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":BF14
            Key             =   "PrintList"
            Object.Tag             =   "622"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":C4AE
            Key             =   "Query"
            Object.Tag             =   "607"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":CA00
            Key             =   "List"
            Object.Tag             =   "608"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":CF52
            Key             =   "DeleteRow"
            Object.Tag             =   "609"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":D4A4
            Key             =   "InsertRow"
            Object.Tag             =   "610"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":D9F6
            Key             =   "AppendRow"
            Object.Tag             =   "611"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":DF48
            Key             =   "Sign"
            Object.Tag             =   "612"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":E49A
            Key             =   "SortUp"
            Object.Tag             =   "613"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":E9EC
            Key             =   "SortDown"
            Object.Tag             =   "614"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":EF3E
            Key             =   "Exit"
            Object.Tag             =   "615"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":F490
            Key             =   "filter"
            Object.Tag             =   "616"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":F9E2
            Key             =   "filter_zot"
            Object.Tag             =   "617"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":FF34
            Key             =   "filter_write"
            Object.Tag             =   "618"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":10486
            Key             =   "filter_clear"
            Object.Tag             =   "619"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":109D8
            Key             =   "Delete"
            Object.Tag             =   "620"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":10AEA
            Key             =   "Execute"
            Object.Tag             =   "621"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImglistBottoniXP 
      Left            =   2100
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   43
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":10BFC
            Key             =   "moduli"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":11196
            Key             =   "Apri"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":11730
            Key             =   "registra"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":11CCA
            Key             =   "annulla"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":12264
            Key             =   "primo"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":127FE
            Key             =   "precedente"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":12D98
            Key             =   "succesivo"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":13332
            Key             =   "ultimo"
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":138CC
            Key             =   "trova"
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":13E66
            Key             =   "rapido"
            Object.Tag             =   "10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":15390
            Key             =   "barragiu"
            Object.Tag             =   "11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":1592A
            Key             =   "barrasu"
            Object.Tag             =   "12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":15EC4
            Key             =   "stampa"
            Object.Tag             =   "13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":1645E
            Key             =   "utente"
            Object.Tag             =   "14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":169F8
            Key             =   "agenti"
            Object.Tag             =   "15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":16B52
            Key             =   "schedulatore"
            Object.Tag             =   "16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":170EC
            Key             =   "inserimento"
            Object.Tag             =   "17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":17686
            Key             =   "help"
            Object.Tag             =   "18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":17C20
            Key             =   "accessi"
            Object.Tag             =   "19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":181BA
            Key             =   "MenuVerticale"
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":1900C
            Key             =   "IconaMXP"
            Object.Tag             =   "21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":193A6
            Key             =   "Zoom"
            Object.Tag             =   "22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":19940
            Key             =   "AllInOne"
            Object.Tag             =   "23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":1FBDA
            Key             =   "Designer"
            Object.Tag             =   "24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":204B4
            Key             =   "New"
            Object.Tag             =   "25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":207CE
            Key             =   "Info"
            Object.Tag             =   "26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":20D68
            Key             =   "dbase"
            Object.Tag             =   "27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":20EC2
            Key             =   "Lingua"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":212C1
            Key             =   "Euro"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":2141B
            Key             =   "NomiCmp"
            Object.Tag             =   "54"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":21575
            Key             =   "SitAnagr"
            Object.Tag             =   "55"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":216CF
            Key             =   "VersDesigner"
            Object.Tag             =   "32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":21C69
            Key             =   "Log"
            Object.Tag             =   "56"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":22203
            Key             =   "RicProfili"
            Object.Tag             =   "57"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":2279D
            Key             =   "VisibVersioni"
            Object.Tag             =   "59"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":22D37
            Key             =   "GrigliaDesign"
            Object.Tag             =   "61"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":232D1
            Key             =   "Messaggi"
            Object.Tag             =   "37"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":2386B
            Key             =   "VisLog"
            Object.Tag             =   "38"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":23E05
            Key             =   "Theme"
            Object.Tag             =   "39"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":2439F
            Key             =   "Prefs"
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":24939
            Key             =   "dettagli"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":24ED3
            Key             =   "OriginalSize"
            Object.Tag             =   "42"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI2005Nucleo.frx":2546D
            Key             =   "ie"
            Object.Tag             =   "43"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar BarraStato1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7215
      Visible         =   0   'False
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11536
            MinWidth        =   847
            Key             =   "info"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Picture         =   "MDI2005Nucleo.frx":256F8
            Key             =   "dittaanno"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ditta Attiva"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Picture         =   "MDI2005Nucleo.frx":258D2
            Key             =   "utente"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Utente Attivo"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   529
            MinWidth        =   529
            Key             =   "save"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Picture         =   "MDI2005Nucleo.frx":25AAC
            Key             =   "Designer"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Versione MyERP"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1032
            MinWidth        =   1027
            Picture         =   "MDI2005Nucleo.frx":25C86
            Key             =   "Lingua"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Lingua Attiva"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "02/10/2012"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   3540
      Top             =   2820
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   2580
      Top             =   3180
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   10
   End
End
Attribute VB_Name = "metodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
DefInt A-Z


Const ANA_AGE = 1  'ultimo=28
Const ANA_ART = 2
Const ANA_ARTVAR = 28
Const ANA_BAN = 3
Const ANA_CLI = 4
Const ANA_DIT = 5
Const ANA_FOR = 6
Const ANA_GEN = 7
Const CAU_MAG = 8
Const TAB_ALI = 9
Const TAB_CAM = 10
Const TAB_CAT = 11
Const TAB_CAU = 12
Const TAB_CCO = 13
Const TAB_ESE = 14
Const TAB_GRU = 15
Const TAB_LIN = 16
Const TAB_MAS = 17
Const TAB_NAZ = 18
Const TAB_PAG = 19
Const TAB_SET = 20
Const TAB_SPE = 21
Const TAB_VAR = 22
Const TAB_ZON = 23
Const VIN_COR = 24
Const VIN_FAT = 25
Const VIN_GEN = 26
Const VIN_VEN = 27


Const ID_POPUPLINGUEBASE = 17000


'=======================
'       variabili
'=======================
Dim mBolLoadingMetodo As Boolean
Dim mBolUnloadingMetodo As Boolean
Dim mIntFormsCount As Integer
Dim mEnmBottoneAgenti As setBottoneAgente
Dim mFunzioniM98 As CFunzioniMetodo98
Dim MEmptyform As Form
Dim MDockAperto As Boolean

Private WithEvents MFormAttiva As Form
Attribute MFormAttiva.VB_VarHelpID = -1
Private WithEvents StatusBar As XtremeCommandBars.StatusBar
Attribute StatusBar.VB_VarHelpID = -1

Public Barra As cBarra
Public BarraStato As CBarraStato

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Dim MLngSplitter As Long
Dim MintIDLinguaMax As Integer
'Utilizzata all'interno di MXDesigner
Public Sub AttivaBottoneDesigner(ByVal TipoBottone As String, ByVal bolAttiva As Boolean)
    Select Case UCase(TipoBottone)
        Case "SETTING"
            FindCommandBar(ID_TLB_DESIGNER).FindControl(, idxDesignerSetting).Enabled = bolAttiva
        Case "VISIBLE"
            FindCommandBar(ID_TLB_DESIGNER).FindControl(, idxDesignerUnVisible).Enabled = bolAttiva
        Case "VISIBLEVALUE"
            FindCommandBar(ID_TLB_DESIGNER).FindControl(, idxDesignerUnVisible).Checked = bolAttiva
        Case "DISABLED"
            FindCommandBar(ID_TLB_DESIGNER).FindControl(, idxDesignerDisable).Enabled = bolAttiva
        Case "DISABLEDVALUE"
            FindCommandBar(ID_TLB_DESIGNER).FindControl(, idxDesignerDisable).Checked = bolAttiva
    End Select
End Sub

Public Sub CambiaFormAttiva(ByVal frmHWnd As Long, frmName As String)
    While IsWindow(frmHWnd)
        DoEvents
        Sleep 10   'Evito CPU 100%
    Wend
    If MFormAttiva.Name = frmName Then
        Set MFormAttiva = Nothing
    End If
End Sub

Friend Sub CambiaColoriFormAttive(Optional RePaint As Boolean = False)
    Dim q As Long
    Dim strNome  As String
    For q = 0 To Forms.Count - 1
        strNome = Forms(q).Name
        If StrComp(strNome, "metodo", vbTextCompare) <> 0 Then
            If Not EsisteElementoCollection(McolFormsInNavBar, strNome) Then
                Call CambiaColoriControlli(Forms(q))
                If RePaint Then
                    If Forms(q).Name <> "frmUserPanelContainer" Then
                        Forms(q).Visible = False
                        Forms(q).Visible = True
                    End If
                End If
            End If
        End If
    Next q

End Sub

Private Sub CambiaLingua(ByVal pLingua As String)
    Dim strLingua As String

    strLingua = Trim(pLingua)
    If MXNU.LinguaAttiva <> strLingua Then
        Call ChiudiFormAttive
        Call MXNU.CambiaRisorse(strLingua)
        
        metodo.MousePointer = vbHourglass
        Call CaricaModuliMetodo(True)
        Call CambiaRisorseCJ
        frmModuli.ModuloAttivo = "Metodo98"
        Call AggiornaStatusBar
        metodo.MousePointer = vbNormal
    End If

End Sub

Friend Sub CaricaLayout()
    Dim strRiga As String, strEntry As String, strValore As String
    Dim strCmdBarsState As String, strDockPaneState As String, strTemaAttivo As String, strSplitter As String, strAnim As String, strPanel As String, strResizer As String, strColorScheme As String
    Dim colRigheIni As New Collection, i&, NomeSkin As String
    Dim f As Integer
    On Local Error Resume Next
    If Dir$(MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini", vbNormal) <> "" Then
        f = FreeFile
        Open MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini" For Input As #f
        
        Do While Not EOF(f)
            Line Input #f, strRiga
            If Trim(strRiga) <> "" Then
                If Left(strRiga, 1) <> ";" Then
                    colRigheIni.Add strRiga
                End If
            End If
        Loop
        
        Close #f
        
        For i = 1 To colRigheIni.Count
            strEntry = Left(colRigheIni(i), InStr(colRigheIni(i), "=") - 1)
            strValore = Mid(colRigheIni(i), InStr(colRigheIni(i), "=") + 1)
            Select Case LCase(strEntry)
                Case "commandbars": strCmdBarsState = strValore
                Case "dockingpane": strDockPaneState = strValore
                Case "temaattivo": strTemaAttivo = strValore
                Case "splitterpreferiti": strSplitter = strValore
                Case "animazionepannelli": strAnim = strValore
                Case "pannellivisibili": strPanel = strValore
                Case "usaresizer": strResizer = strValore
                Case "schemacolori": strColorScheme = strValore
            End Select
        Next
        If strTemaAttivo = "" Then strTemaAttivo = "Office2007"
        If strAnim = "" Then strAnim = "1"
        
        'rif. anomalia 9469
        For i = ID_TLBITEM_THEME_OFFICE2007 To ID_TLBITEM_THEME_VISTA
            FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, i).DefaultItem = False
        Next i
        
        'FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, btnId).DefaultItem = True
        
        
        Select Case LCase(strTemaAttivo)
            Case "office2007"
                'SkinFramework.LoadSkin MXNU.PercorsoPgm & "\Themes\Office2007", ""
                NomeSkin = "Office2007"
                FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2007).DefaultItem = True
            Case "winxpluna"
                'SkinFramework.LoadSkin MXNU.PercorsoPgm & "\Themes\WinXP.Luna", ""
                NomeSkin = "WinXP.Luna"
                FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_WINXPLUNA).DefaultItem = True
            Case "winxproyale"
                'SkinFramework.LoadSkin MXNU.PercorsoPgm & "\Themes\WinXP.Royale", ""
                NomeSkin = "WinXP.Royale"
                FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_WINXPROYALE).DefaultItem = True
            Case "vista"
                'SkinFramework.LoadSkin MXNU.PercorsoPgm & "\Themes\Vista", ""
                NomeSkin = "Vista"
                FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_VISTA).DefaultItem = True
            Case "system"
                'SkinFramework.LoadSkin MXNU.PercorsoPgm & "", ""
                NomeSkin = ""
                FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_SYSTEM).DefaultItem = True
            Case Else
                'SkinFramework.LoadSkin MXNU.PercorsoPgm & "\Themes\Office2007", ""
                NomeSkin = "Office2007"
                FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2007).DefaultItem = True
        End Select
        'MXCtrl.SkinnedTheme = (strTemaAttivo <> "" And LCase(strTemaAttivo) <> "system")
        'SkinFramework.EnableThemeDialogTexture Me.hwnd, 6
        'SkinFramework.ApplyWindow Me.hwnd
        'SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
        'Call frmOpzioniGenerali.CambiaColorScheme(0)
        Call CambiaTema(NomeSkin)

        DoEvents
        If strCmdBarsState <> "" Then CommandBars.LoadStateFromString strCmdBarsState, True, True, True
        If strDockPaneState <> "" Then DockingPaneManager.LoadStateFromString strDockPaneState
        MDockAperto = Not DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden
        DockingPaneManager.FindPane(ID_PANE_NAVBAR).Select


        'verifico la presenza o meno della messaggistica
        '-esiste: rigenero il pane della messaggistica
        '-nonesiste: non faccio nulla
        
        If (Not mMessagingEngine Is Nothing) Then
            If DockingPaneManager.FindPane(ID_PANE_TASKBAR2) Is Nothing Then
                Dim PaneTaskBar2 As Pane
                Set PaneTaskBar2 = DockingPaneManager.CreatePane(ID_PANE_TASKBAR2, 250, 200, DockBottomOf, Nothing)
                With PaneTaskBar2
                    .Title = MXNU.CaricaStringaRes(11991)  '"Messaggi"
                    '.AttachTo PaneTaskBar1
                    .IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "Messaggi")
                End With
                PaneTaskBar2.Hide
            End If
        Else
            If Not DockingPaneManager.FindPane(ID_PANE_TASKBAR2) Is Nothing Then
               Call DockingPaneManager.DestroyPane(DockingPaneManager.FindPane(ID_PANE_TASKBAR2))
            End If
        End If
On Local Error Resume Next
        Call AppActivate(Me.Caption)
On Local Error GoTo 0

        If Cast(strAnim, vbInteger) = 1 Then
            metodo.DockingPaneManager.Options.AnimationDelay = -1
        Else
            metodo.DockingPaneManager.Options.AnimationDelay = 0
        End If
        If strPanel <> "" Then
            frmModuli2005.ShortcutBar1.ExpandedLinesCount = Cast(strPanel, vbInteger)
        End If
        If strResizer = "" Then
            MXNU.UsaResizer = True
        Else
            MXNU.UsaResizer = (strResizer = "1")
        End If
    Else
        'Se per l'utente attivo non c'è un file di layout, imposto Office2010 come tema attivo
        'MXCtrl.SkinnedTheme = True
        Call CambiaTema("Office2010")
        FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2007).DefaultItem = True
        MXNU.UsaResizer = True   'Per default il resizer su Evolus è attivo
    End If
'    If CommandBars.VisualTheme = xtpThemeOffice2003 Then
'        Select Case LCase(strColorScheme)
'            Case "", "default": Call frmOpzioniGenerali.CambiaColorScheme(0)
'            Case "blue": Call frmOpzioniGenerali.CambiaColorScheme(1)
'            Case "silver": Call frmOpzioniGenerali.CambiaColorScheme(2)
'            Case "olive": Call frmOpzioniGenerali.CambiaColorScheme(3)
'        End Select
'        Call CambiaTema(xtpThemeOffice2003)
'    End If
    
    'Nascondo le eventuali form dei pannelli utente personalizzati, altrimenti le estensioni col Wrapper si vedono "fluttuanti" in mezzo alla navbar
    For i = 0 To Forms.Count - 1
        If StrComp(Forms(i).Name, "frmUserPanelContainer", vbTextCompare) = 0 Then
            Forms(i).Hide
        End If
    Next i
    'Ripristino il titolo del docking pane della MDI in quanto alla partenza di Metodo è sempre attiva la sezione Moduli Programma
    DockingPaneManager.FindPane(ID_PANE_NAVBAR).Title = frmModuli2005.ShortcutBar1.FindItem(ID_BAR_PROGMODULES).Caption
    
    'Nascondo le toolbar di MyErp e Quality, qualora si sia usciti da Metodo con la toolbar attiva.
    metodo.FindCommandBar(ID_TLB_DESIGNER).Visible = False
    metodo.FindCommandBar(ID_TLB_QUALITY).Visible = False
        
    'Reimposto le risorse in lingua per la CodeJock, qualora si sia usciti da Metodo con una lingua diversa da quella di default
    Call CambiaRisorseCJ

    On Local Error Resume Next
    'Rimuovo la toolbar del browser se la classe corrispondente non è stata inizializzata (il carica layout la visualizza lo stesso se era stata salvata in precedenza uscendo da metodo)
    If (mMetodoBrowser Is Nothing) Then
        Dim tlbBrowser As XtremeCommandBars.CommandBar
        Set tlbBrowser = FindCommandBar(ID_TLB_BROWSER)
        If Not (tlbBrowser Is Nothing) Then
            tlbBrowser.Delete
        End If
        Set tlbBrowser = Nothing
    End If
    
'    MXCtrl.SkinnedTheme = (NomeFileTema <> "")
'    SkinFramework.LoadSkin MXNU.PercorsoPgm & "\Themes\Office2007", ""
'    SkinFramework.EnableThemeDialogTexture Me.hwnd, 6
'    SkinFramework.ApplyWindow Me.hwnd
'    SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
    If strSplitter <> "" Then
        'frmModuli.MWSplitter1.SplitterTop = Cast(strSplitter, vbLong)
        MLngSplitter = Cast(strSplitter, vbLong)
    End If
    On Local Error GoTo 0
    

End Sub

'Carica posizione, dimensione e stato di Toolbar, MenuBar e DockingPane
Friend Sub CaricaLayout_OLD()
    Dim strRiga As String, strEntry As String, strValore As String
    Dim strCmdBarsState As String, strDockPaneState As String, strTemaAttivo As String, strSplitter As String, strAnim As String, strPanel As String, strResizer As String, strColorScheme As String
    Dim colRigheIni As New Collection, i&
    Dim f As Integer
    On Local Error Resume Next
    If Dir$(MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini", vbNormal) <> "" Then
        f = FreeFile
        Open MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini" For Input As #f
        
        Do While Not EOF(f)
            Line Input #f, strRiga
            If Trim(strRiga) <> "" Then
                If Left(strRiga, 1) <> ";" Then
                    colRigheIni.Add strRiga
                End If
            End If
        Loop
        
        Close #f
        
        For i = 1 To colRigheIni.Count
            strEntry = Left(colRigheIni(i), InStr(colRigheIni(i), "=") - 1)
            strValore = Mid(colRigheIni(i), InStr(colRigheIni(i), "=") + 1)
            Select Case LCase(strEntry)
                Case "commandbars": strCmdBarsState = strValore
                Case "dockingpane": strDockPaneState = strValore
                Case "temaattivo": strTemaAttivo = strValore
                Case "splitterpreferiti": strSplitter = strValore
                Case "animazionepannelli": strAnim = strValore
                Case "pannellivisibili": strPanel = strValore
                Case "usaresizer": strResizer = strValore
                Case "schemacolori": strColorScheme = strValore
            End Select
        Next
        If strTemaAttivo = "" Then strTemaAttivo = "Office2003"
        If strAnim = "" Then strAnim = "1"
        
        If strCmdBarsState <> "" Then CommandBars.LoadStateFromString strCmdBarsState, True, True, True
        If strDockPaneState <> "" Then DockingPaneManager.LoadStateFromString strDockPaneState
        Select Case LCase(strTemaAttivo)
            Case "office2000"
                Call CambiaTema(xtpThemeOffice2000)
                '*FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2000).DefaultItem = True
            Case "officexp"
                Call CambiaTema(xtpThemeOfficeXP)
                '*FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICEXP).DefaultItem = True
            Case "office2003"
                Call CambiaTema(xtpThemeOffice2003)
                '*FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2003).DefaultItem = True
            Case "xpnative"
                Call CambiaTema(xtpThemeNativeWinXP)
                '*FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_NATIVEXP).DefaultItem = True
        End Select
        'Spostato dopo la visualizzazione della form e dopo l'inizializzazione dello SkinFramework - Vedi MDIForm_Load
'        If strSplitter <> "" Then
'            frmModuli.MWSplitter1.SplitterTop = Cast(strSplitter, vbLong)
'        End If
        If Cast(strAnim, vbInteger) = 1 Then
            metodo.DockingPaneManager.Options.AnimationDelay = -1
        Else
            metodo.DockingPaneManager.Options.AnimationDelay = 0
        End If
        If strPanel <> "" Then
            frmModuli2005.ShortcutBar1.ExpandedLinesCount = Cast(strPanel, vbInteger)
        End If
        If strResizer = "" Then
            MXNU.UsaResizer = True
        Else
            MXNU.UsaResizer = (strResizer = "1")
        End If
    Else
        'Se per l'utente attivo non c'è un file di layout, imposto Office2003 come tema attivo
        Call CambiaTema(xtpThemeOffice2003)
        '*FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2003).DefaultItem = True
        MXNU.UsaResizer = True   'Per default il resizer su Evolus è attivo
    End If
    
    'Nascondo le eventuali form dei pannelli utente personalizzati, altrimenti le estensioni col Wrapper si vedono "fluttuanti" in mezzo alla navbar
    For i = 0 To Forms.Count - 1
        If StrComp(Forms(i).Name, "frmUserPanelContainer", vbTextCompare) = 0 Then
            Forms(i).Hide
        End If
    Next i
    'Ripristino il titolo del docking pane della MDI in quanto alla partenza di Metodo è sempre attiva la sezione Moduli Programma
    DockingPaneManager.FindPane(ID_PANE_NAVBAR).Title = frmModuli2005.ShortcutBar1.FindItem(ID_BAR_PROGMODULES).Caption
    
    'Nascondo le toolbar di MyErp e Quality, qualora si sia usciti da Metodo con la toolbar attiva.
    metodo.FindCommandBar(ID_TLB_DESIGNER).Visible = False
    metodo.FindCommandBar(ID_TLB_QUALITY).Visible = False
        
    'Reimposto le risorse in lingua per la CodeJock, qualora si sia usciti da Metodo con una lingua diversa da quella di default
    Call CambiaRisorseCJ
    
    On Local Error Resume Next
    'Rimuovo la toolbar del browser se la classe corrispondente non è stata inizializzata (il carica layout la visualizza lo stesso se era stata salvata in precedenza uscendo da metodo)
    If (mMetodoBrowser Is Nothing) Then
        Dim tlbBrowser As XtremeCommandBars.CommandBar
        Set tlbBrowser = FindCommandBar(ID_TLB_BROWSER)
        If Not (tlbBrowser Is Nothing) Then
            tlbBrowser.Delete
        End If
        Set tlbBrowser = Nothing
    End If
    
    If strSplitter <> "" Then
        'frmModuli.MWSplitter1.SplitterTop = Cast(strSplitter, vbLong)
        MLngSplitter = Cast(strSplitter, vbLong)
    End If
    
    On Local Error GoTo 0
    
End Sub

Private Sub CreatePanes()
    Dim PaneOff2003 As Pane
    Dim PaneTaskBar1 As Pane, PaneTaskBar2 As Pane
    
    'Important!!!
    DockingPaneManager.SetCommandBars metodo.CommandBars
    
    'Risorse in lingua per i componenti Codejock
    DockingPaneGlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"
    
    'Set PaneTaskBar1 = DockingPaneManager.CreatePane(ID_PANE_TASKBAR1, 250, 200, DockBottomOf, Nothing)
    'With PaneTaskBar1
    '    .Title = "Log"
    '    .IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "VisLog")
    '    .Options = PaneNoCloseable
    'End With
    'Set PaneTaskBar2 = DockingPaneManager.CreatePane(ID_PANE_TASKBAR2, 250, 200, DockLeftOf, PaneTaskBar1)
    If (Not mMessagingEngine Is Nothing) Then
        Set PaneTaskBar2 = DockingPaneManager.CreatePane(ID_PANE_TASKBAR2, 250, 200, DockBottomOf, Nothing)
        With PaneTaskBar2
            .Title = MXNU.CaricaStringaRes(11991)  '"Messaggi"
            '.AttachTo PaneTaskBar1
            .IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "Messaggi")
        End With
        PaneTaskBar2.Hide
    End If
    
'    'pane per metodo browser
'    If (Not mMetodoBrowser Is Nothing) Then
'        Dim browserPane As Pane
'        Set browserPane = mMetodoBrowser.CreatePane(Me.DockingPaneManager)
'        browserPane.Hide
'    End If

    'PaneTaskBar1.Hide
    Set PaneOff2003 = DockingPaneManager.CreatePane(ID_PANE_NAVBAR, 250, 120, DockLeftOf, Nothing)
    PaneOff2003.Title = MXNU.CaricaStringaRes(11989)  '"Moduli Programma"
    PaneOff2003.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "rapido")
    PaneOff2003.Close
    PaneOff2003.Options = PaneNoCloseable
    'PaneOff2003.Hide
    
    DockingPaneManager.PanelPaintManager.Layout = xtpTabLayoutAutoSize
    DockingPaneManager.Options.ThemedFloatingFrames = True
    DockingPaneManager.ImageList = ImglistBottoniXP
    
    
End Sub

Private Sub CreateStatusBar()
    Dim Pane As StatusBarPane
    Set StatusBar = CommandBars.StatusBar
    StatusBar.Visible = True
    
    Set Pane = StatusBar.AddPane(STATUSBAR_ID_PANELINFO)
    Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
    Pane.Text = ""
    Pane.Width = 0  'AutoSize
    
    Set Pane = StatusBar.AddPane(STATUSBAR_ID_PANELDITTAANNO)
    Pane.ToolTip = "Ditta Attiva"
    Pane.IconIndex = metodo.ImglistBottoniXP.ListImages("dbase").Index
    Pane.BeginGroup = True
    Pane.Width = 2000 \ Screen.TwipsPerPixelX
    
    
    Set Pane = StatusBar.AddPane(STATUSBAR_ID_PANELUTENTE)
    Pane.ToolTip = "Utente Attivo"
    Pane.IconIndex = metodo.ImglistBottoniXP.ListImages("utente").Index
    Pane.BeginGroup = True
    Pane.Width = 1440 \ Screen.TwipsPerPixelX
    
    Set Pane = StatusBar.AddPane(STATUSBAR_ID_PANELSAVE)
    Pane.ToolTip = "Transazioni Attive"
    Pane.IconIndex = metodo.ImglistBottoniXP.ListImages("registra").Index
    Pane.BeginGroup = True
    Pane.Width = 300 \ Screen.TwipsPerPixelX
    Pane.Visible = False
    
    Set Pane = StatusBar.AddPane(STATUSBAR_ID_PANELDESIGNER)
    Pane.ToolTip = "Versione MyERP"
    Pane.IconIndex = metodo.ImglistBottoniXP.ListImages("VersDesigner").Index
    Pane.Button = True
    Pane.BeginGroup = True
    Pane.Width = 2000 \ Screen.TwipsPerPixelX
    
    
    Set Pane = StatusBar.AddPane(STATUSBAR_ID_PANELLINGUA)
    Pane.ToolTip = "Lingua Attiva"
    Pane.Button = True
    Pane.BeginGroup = True
    Pane.Width = 1440 \ Screen.TwipsPerPixelX
    
    StatusBar.RibbonDividerIndex = 0
    StatusBar.ToolTipContext.Style = xtpToolTipOffice
    StatusBar.EnableCustomization False
    StatusBar.IdleText = ""
End Sub

'Creazione delle Toolbar per la MDI di Metodo.
'ATTENZIONE: L'Id di ogni singolo bottone DEVE ESSERE UNIVOCO rispetto a tutti i controlli inseriti nelle varie toolbar/menu altrimenti poi non si sa che controllo è stato selezionato
'ATTENZIONE: Se si agiungono bottoni e si modificano gli enum collegati ricordarsi di modificare anche l'evento CommandBars_Execute
Public Sub CreateToolBars()
    Dim tlbPrinc As XtremeCommandBars.CommandBar
    Dim tlbQuality As XtremeCommandBars.CommandBar
    Dim tlbDesigner As XtremeCommandBars.CommandBar
    Dim Btn As CommandBarControl
    Dim btnPopup As CommandBarPopup
    
    'CommandBars.GlobalSettings.App = App
    CommandBarsGlobalSettings.App = App
    CommandBars.EnableCustomization False
    'Risorse in lingua per i componenti Codejock
    'CommandBars.GlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"
    CommandBarsGlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"
    
        
    'ToolBar Principale
    Set tlbPrinc = FindCommandBar(ID_TLB_PRINC)
    If tlbPrinc Is Nothing Then
        Set tlbPrinc = CommandBars.Add("Barra", xtpBarTop)
        tlbPrinc.BarID = ID_TLB_PRINC
        tlbPrinc.ContextMenuPresent = False
    End If
    tlbPrinc.Closeable = False
    tlbPrinc.Controls.DeleteAll
    CommandBars.Options.KeyboardCuesUse = xtpKeyboardCuesUseNone   'Anomalia 8691
    CommandBars.KeyBindings.DeleteAll   'Necessario altrimenti in caso di caricamento di layout vecchi vengono reimpostate le combinazioni di tasti vecchie (es. F12 per aprire le Opzioni Generali)
    With tlbPrinc.Controls
        Set Btn = .Add(xtpControlButton, idxBottoneModuli, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "moduli")
        Btn.BeginGroup = True
        Btn.Checked = True
        Btn.ToolTipText = MXNU.CaricaStringaRes(1)
        Set Btn = .Add(xtpControlButton, idxBottoneInserisci, "")
        Btn.BeginGroup = True
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "inserimento")
        Btn.ToolTipText = MXNU.CaricaStringaRes(2)
        Set Btn = .Add(xtpControlButton, idxBottoneDettaglio, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "dettagli")
        Btn.ToolTipText = MXNU.CaricaStringaRes(3)
        Set Btn = .Add(xtpControlButton, idxBottoneRegistra, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "registra")
        Btn.ToolTipText = MXNU.CaricaStringaRes(4)
        Set Btn = .Add(xtpControlButton, idxBottoneAnnulla, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "annulla")
        Btn.ToolTipText = MXNU.CaricaStringaRes(5)
        Set Btn = .Add(xtpControlButton, idxBottonePrimo, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "primo")
        Btn.ToolTipText = MXNU.CaricaStringaRes(6)
        Btn.BeginGroup = True
        Set Btn = .Add(xtpControlButton, idxBottonePrecedente, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "precedente")
        Btn.ToolTipText = MXNU.CaricaStringaRes(7)
        Set Btn = .Add(xtpControlButton, idxBottoneSuccessivo, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "succesivo")
        Btn.ToolTipText = MXNU.CaricaStringaRes(8)
        Set Btn = .Add(xtpControlButton, idxBottoneUltimo, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "ultimo")
        Btn.ToolTipText = MXNU.CaricaStringaRes(9)
        Set Btn = .Add(xtpControlButton, idxBottoneTrova, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "trova")
        Btn.ToolTipText = MXNU.CaricaStringaRes(10)
        Btn.BeginGroup = True
        
'**** Bottone Zoom tolto su 8.04.00 *******************************************************
'        'Bottone Zoom
'        Set btnPopup = .Add(xtpControlSplitButtonPopup, idxBottoneZoom, "")
'        btnPopup.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "Zoom")
'        btnPopup.ToolTipText = MXNU.CaricaStringaRes(77)
'        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneZoom100, "100%")
'        Btn.IconId = 10000000
'        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneZoom150, "150%")
'        Btn.IconId = 10000000
'        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneZoom200, "200%")
'        Btn.IconId = 10000000
'******************************************************************************************
                
        Set Btn = .Add(xtpControlButton, idxBottoneStampa, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "stampa")
        Btn.ToolTipText = MXNU.CaricaStringaRes(11)
        Btn.BeginGroup = True
        
        Set Btn = .Add(xtpControlLabel, ID_TLBITEM_SPACER1, Space(15))
        Btn.ID = ID_TLBITEM_SPACER1   '<<< Riassegno l'ID altrimenti è 0 (?!?)
        Set Btn = .Add(xtpControlLabel, ID_TLBITEM_LBLMODULO, "Metodo Evolus")
        Btn.ID = ID_TLBITEM_LBLMODULO   '<<< Riassegno l'ID altrimenti è 0 (?!?)
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "IconaMXP")
        Set Btn = .Add(xtpControlLabel, ID_TLBITEM_SPACER2, Space(17))
        Btn.ID = ID_TLBITEM_SPACER2   '<<< Riassegno l'ID altrimenti è 0 (?!?)
        
        Set Btn = .Add(xtpControlButton, idxBottoneAllInOne, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "AllInOne")
        Btn.ToolTipText = MXNU.CaricaStringaRes(76)
        
        'Bottone Designer
        Set btnPopup = .Add(xtpControlSplitButtonPopup, idxBottoneDesigner, "")
        btnPopup.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "Designer")
        btnPopup.ToolTipText = MXNU.CaricaStringaRes(78)
        btnPopup.BeginGroup = True
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneAttivaDesigner, MXNU.CaricaStringaRes(80))
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "Designer")
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottonePrioritaVisibVers, MXNU.CaricaStringaRes(79))
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "VisibVersioni")
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneCancVersione, MXNU.CaricaStringaRes(89))
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "annulla")
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneGrigliaDesigner, MXNU.CaricaStringaRes(91))
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "GrigliaDesign")
        
        Set Btn = .Add(xtpControlButton, ID_TLBITEM_FRMORIGINALSIZE, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "OriginalSize")
        Btn.ToolTipText = MXNU.CaricaStringaRes(11994)
                        
        
        Set Btn = .Add(xtpControlButton, idxBottoneUtenteModifica, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "utente")
        Btn.ToolTipText = MXNU.CaricaStringaRes(12)
        Btn.BeginGroup = True
        
        ' S#3040 - rimossa la Gestione Accessi da Evolus
'        Set Btn = .Add(xtpControlButton, idxBottoneDefAccessi, "")
'        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "accessi")
'        Btn.ToolTipText = MXNU.CaricaStringaRes(13)

        
        Set Btn = .Add(xtpControlButton, idxBottoneDefAgenti, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "agenti")
        Btn.ToolTipText = MXNU.CaricaStringaRes(14)
        Btn.BeginGroup = True
        Set Btn = .Add(xtpControlButton, idxBottoneNomiCtrlCmp, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "NomiCmp")
        Btn.ToolTipText = MXNU.CaricaStringaRes(19)
        Set Btn = .Add(xtpControlButton, idxBottoneSituazAnagr, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "SitAnagr")
        Btn.ToolTipText = MXNU.CaricaStringaRes(20)
        Set Btn = .Add(xtpControlButton, idxBottoneAttivaFileLog, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "Log")
        Btn.ToolTipText = MXNU.CaricaStringaRes(21)
        Set Btn = .Add(xtpControlButton, idxBottoneRicProfili, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "RicProfili")
        Btn.ToolTipText = MXNU.CaricaStringaRes(22)

        'Bottone Agenti
'        Set btnPopup = .Add(xtpControlSplitButtonPopup, idxBottoneDefAgenti, "")
'        btnPopup.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "agenti")
'        btnPopup.ToolTipText = MXNU.CaricaStringaRes(14)
'        Set btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneAttivaDefAgenti, MXNU.CaricaStringaRes(18))
'        btn.IconId = 10000000
'        Set btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneNomiCtrlCmp, MXNU.CaricaStringaRes(19))
'        btn.IconId = 10000000
'        Set btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneSituazAnagr, MXNU.CaricaStringaRes(20))
'        btn.IconId = 10000000
'        Set btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneAttivaFileLog, MXNU.CaricaStringaRes(21))
'        btn.IconId = 10000000
'        Set btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxBottoneRicProfili, MXNU.CaricaStringaRes(22))
'        btn.IconId = 10000000
        
        
        Set Btn = .Add(xtpControlButton, idxBottoneHelp, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "help")
        Btn.ToolTipText = MXNU.CaricaStringaRes(15)
        Btn.BeginGroup = True
        Set Btn = .Add(xtpControlButton, idxBottoneSchedula, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "schedulatore")
        Btn.ToolTipText = MXNU.CaricaStringaRes(16)
        
        Set Btn = .Add(xtpControlButton, ID_TLBITEM_INFO, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "Info")
        Btn.ToolTipText = MXNU.CaricaStringaRes(92)
        Btn.BeginGroup = True
        
        Set Btn = .Add(xtpControlButton, ID_TLBITEM_PREFERENCES, "")
        Btn.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "Prefs")
        Btn.ToolTipText = MXNU.CaricaStringaRes(11908) & " (Ctrl+O)"
        Btn.BeginGroup = True
        'CommandBars.KeyBindings.AllowDoubleKeyShortcuts = True
        'CommandBars.KeyBindings.Add 0, vbKeyF12, ID_TLBITEM_PREFERENCES
        CommandBars.KeyBindings.Add 8, vbKeyO, ID_TLBITEM_PREFERENCES    'FCONTROL = 8
        
        'Bottone Cambio Tema
        Set btnPopup = .Add(xtpControlSplitButtonPopup, ID_TLBITEM_CHANGETHEME, MXNU.CaricaStringaRes(100))   '"Tema Visuale"
        btnPopup.IconId = ImgListKey2ImgListIdx(ImglistBottoniXP, "Theme")
        'btnPopup.ToolTipText = "Cambia Tema Visuale"
        btnPopup.BeginGroup = True
        btnPopup.Style = xtpButtonIconAndCaption
        
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, ID_TLBITEM_THEME_OFFICE2010, "Office 2010")
        Btn.IconId = 100000
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, ID_TLBITEM_THEME_OFFICE2007, "Office 2007")
        Btn.IconId = 100000
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, ID_TLBITEM_THEME_WINXPLUNA, "Windows XP Luna")
        Btn.IconId = 100000
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, ID_TLBITEM_THEME_WINXPROYALE, "Windows XP Royale")
        Btn.IconId = 100000
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, ID_TLBITEM_THEME_VISTA, "Windows Vista")
        Btn.IconId = 100000
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, ID_TLBITEM_THEME_SYSTEM, "Predefinito di Sistema")
        Btn.IconId = 100000
        
        CommandBars.KeyBindings.Add FCONTROL, Asc("L"), idxBottoneModuli
    End With
    CommandBars.AddImageList ImglistBottoniXP
    
'    If (Screen.Width \ Screen.TwipsPerPixelX) = 800 And (Screen.Height \ Screen.TwipsPerPixelY) = 600 Then
'        .Item(18).Width = 3300 '4350
'        pnlBack.Width = 2350 '2640
'    End If
                
'        'imposto il back color standard xp
'        Me.BackColor = RGB(161, 161, 146)
        
        ' rif. anomalia n.ro 4824 (la voce di menù è attivata solo se ho runtime agenti + sono amministratore)
        tlbPrinc.FindControl(, idxBottoneDefAgenti).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
        tlbPrinc.FindControl(, idxBottoneNomiCtrlCmp).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
        tlbPrinc.FindControl(, idxBottoneSituazAnagr).Enabled = Not (MXNU.CtrlAccessi)
        ' FINE rif. anomalia n.ro 4824
    
    If Not ((MXNU.MetodoXP) _
            And ((MXNU.ControlloModuliChiave(modAllInOneRuntime) = 0) _
            Or (MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0))) Then
            
        tlbPrinc.FindControl(, idxBottoneAllInOne).Visible = False
        tlbPrinc.FindControl(, idxBottoneDesigner).BeginGroup = False 'Barra separatrice
    End If
    
    tlbPrinc.FindControl(, idxBottoneSchedula).Enabled = (MXNU.ControlloModuliChiave(modJobScheduler) = 0)
    
    
    'ToolBar Quality
    Set tlbQuality = FindCommandBar(ID_TLB_QUALITY)
    If tlbQuality Is Nothing Then
        Set tlbQuality = CommandBars.Add("Quality", xtpBarTop)
        tlbQuality.BarID = ID_TLB_QUALITY
        tlbQuality.ContextMenuPresent = False
    End If
    CommandBars.AddImageList imglQuality
    tlbQuality.Closeable = False
    tlbQuality.Controls.DeleteAll
    With tlbQuality.Controls
        Set btnPopup = .Add(xtpControlSplitButtonPopup, idxQualityNew, "")
        btnPopup.ToolTipText = "Nuovo Documento"
        Set btnPopup = .Add(xtpControlSplitButtonPopup, idxQualityOpen, "")
        btnPopup.ToolTipText = "Elenco Documenti"
        
        Set Btn = .Add(xtpControlButton, idxQualitySave, "")
        Btn.ToolTipText = "Salva Documento"
        Set Btn = .Add(xtpControlButton, idxQualityDelete, "")
        Btn.ToolTipText = "Elimina"
        Set Btn = .Add(xtpControlButton, idxQualityPreview, "")
        Btn.ToolTipText = "Anteprima di Stampa"
        Set Btn = .Add(xtpControlButton, idxQualityPrint, "")
        Btn.ToolTipText = "Stampa"
        Set Btn = .Add(xtpControlButton, idxQualityPrintList, "")
        Btn.ToolTipText = "Stampa Elenco"
        Set Btn = .Add(xtpControlButton, idxQualityQuery, "")
        Btn.ToolTipText = "Interrogazioni"
        Set Btn = .Add(xtpControlButton, idxQualityList, "")
        Btn.ToolTipText = "Elenco"
        Set Btn = .Add(xtpControlButton, idxQualityDeleteRow, "")
        Btn.ToolTipText = "Elimina Riga"
        Set Btn = .Add(xtpControlButton, idxQualityInsertRow, "")
        Btn.ToolTipText = "Inserisci Riga"
        Set Btn = .Add(xtpControlButton, idxQualityAppendRow, "")
        Btn.ToolTipText = "Accoda Riga"
        Set Btn = .Add(xtpControlButton, idxQualitySign, "")
        Btn.ToolTipText = "Firma"
        Set Btn = .Add(xtpControlButton, idxQualitySortUp, "")
        Btn.ToolTipText = "Ordinamento Crescente"
        Set Btn = .Add(xtpControlButton, idxQualitySortDown, "")
        Btn.ToolTipText = "Ordinamento Decrescente"
        Set Btn = .Add(xtpControlButton, idxQualityFilterWrite, "")
        Btn.ToolTipText = "Scrivi Filtro"
        Set Btn = .Add(xtpControlButton, idxQualityFilterZot, "")
        Btn.ToolTipText = "Filtro in base a selezione"
        Set Btn = .Add(xtpControlButton, idxQualityFilterClear, "")
        Btn.ToolTipText = "Rimuovi Filtro"
        Set Btn = .Add(xtpControlButton, idxQualityExecute, "")
        Btn.ToolTipText = "Esegui"
        'Set Btn = .Add(xtpControlButton, idxQualityExit, "") rif. anomalia #71
        'Btn.ToolTipText = "Uscita dal programma" rif anomalia #71
        
    End With

    'ToolBar Designer
    Set tlbDesigner = FindCommandBar(ID_TLB_DESIGNER)
    If tlbDesigner Is Nothing Then
        Set tlbDesigner = CommandBars.Add("Designer", xtpBarTop)
        tlbDesigner.BarID = ID_TLB_DESIGNER
        tlbDesigner.ContextMenuPresent = False
    End If
    CommandBars.AddImageList ImmaginiDesigner
    tlbDesigner.Closeable = False
    tlbDesigner.Controls.DeleteAll
    With tlbDesigner.Controls
        Set Btn = .Add(xtpControlComboBox, idxDesignerCmbDesignType, "")
        Btn.BeginGroup = True
        Btn.Width = 129
        Btn.DropDownListStyle = False
        Btn.AddItem MXNU.CaricaStringaRes(75474)   'Controlli
        Btn.AddItem MXNU.CaricaStringaRes(75475)   'Linguette
        
        Set Btn = .Add(xtpControlButton, idxDesignerUnVisible, "")
        Btn.ToolTipText = MXNU.CaricaStringaRes(81)
        Btn.Checked = False
        Set Btn = .Add(xtpControlButton, idxDesignerDisable, "")
        Btn.ToolTipText = MXNU.CaricaStringaRes(82)
        Btn.Checked = False
        
        'Bottone Salva
        Set btnPopup = .Add(xtpControlSplitButtonPopup, idxDesignerSave, "")
        btnPopup.BeginGroup = True
        btnPopup.ToolTipText = MXNU.CaricaStringaRes(85)
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxDesignerSaveDesign, "Salva corrente")
        Btn.IconId = 10000000
        Set Btn = btnPopup.CommandBar.Controls.Add(xtpControlButton, idxDesignerSaveVersion, "Salva Impostazione")
        Btn.IconId = 10000000
        
        Set Btn = .Add(xtpControlButton, idxDesignerDiscard, "")
        Btn.ToolTipText = MXNU.CaricaStringaRes(86)
        Set Btn = .Add(xtpControlButton, idxDesignerShowDiff, "")
        Btn.ToolTipText = MXNU.CaricaStringaRes(87)
        Btn.BeginGroup = True
        Set Btn = .Add(xtpControlButton, idxDesignerUndoCtrl, "")
        Btn.ToolTipText = MXNU.CaricaStringaRes(90)
        Set Btn = .Add(xtpControlButton, idxDesignerSetting, "")
        Btn.ToolTipText = MXNU.CaricaStringaRes(88)
        Btn.BeginGroup = True
        
        Set Btn = .Add(xtpControlLabel, idxDesignerLabelSpacer1, "               ")
        Btn.ID = idxDesignerLabelSpacer1
        Btn.Visible = False
        Set Btn = .Add(xtpControlLabel, idxDesignerLabelTop, "Top: ")
        Btn.ID = idxDesignerLabelTop
        Btn.Visible = False
        Set Btn = .Add(xtpControlLabel, idxDesignerLabelLeft, "Left: ")
        Btn.ID = idxDesignerLabelLeft
        Btn.Visible = False
        Set Btn = .Add(xtpControlLabel, idxDesignerLabelHeight, "Height: ")
        Btn.ID = idxDesignerLabelHeight
        Btn.Visible = False
        Set Btn = .Add(xtpControlLabel, idxDesignerLabelWidth, "Width: ")
        Btn.ID = idxDesignerLabelWidth
        Btn.Visible = False
        Set Btn = .Add(xtpControlLabel, idxDesignerLabelNoVisible, "Non Visibile")
        Btn.ID = idxDesignerLabelNoVisible
        Btn.IconId = ImgListKey2ImgListIdx(ImmaginiDesigner, "NoVisible")
        Btn.Visible = False
        Set Btn = .Add(xtpControlLabel, idxDesignerLabelNoEditable, "Non Editabile")
        Btn.ID = idxDesignerLabelNoEditable
        Btn.IconId = ImgListKey2ImgListIdx(ImmaginiDesigner, "NoEditable")
        Btn.Visible = False
        Set Btn = .Add(xtpControlLabel, idxDesignerLabelPosAltered, "Posizione Alterata")
        Btn.ID = idxDesignerLabelPosAltered
        Btn.IconId = ImgListKey2ImgListIdx(ImmaginiDesigner, "PosAltered")
        Btn.Visible = False
        Set Btn = .Add(xtpControlLabel, idxDesignerLabelTabAltered, "Tab Alterato")
        Btn.ID = idxDesignerLabelTabAltered
        Btn.IconId = ImgListKey2ImgListIdx(ImmaginiDesigner, "TabAltered")
        Btn.Visible = False
    End With

    'toolbar browser
    If (Not mMetodoBrowser Is Nothing) Then
        Dim tlbBrowser As XtremeCommandBars.CommandBar
        Set tlbBrowser = mMetodoBrowser.CreateToolBar(Me)   'CommandBars
    End If

    Dim strEntryMw As String
    Dim bolDesignerAbilitato As Boolean
    bolDesignerAbilitato = True
    strEntryMw = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MW.INI", "METODOW", "ATTIVADESIGNER", "")
    If (Len(strEntryMw) > 0) Then
        bolDesignerAbilitato = (StrComp(strEntryMw, "DESIGNER", vbTextCompare) = 0)
    End If
    If Not (bolDesignerAbilitato _
        And MXNU.MetodoXP _
        And MXNU.ControlloModuliChiave(modMyErpDesign) = 0 _
        And Not (MXNU.CtrlAccessi)) Then
        
        'DESIGNER : Rendo invisibile il bottone se il Designer di MyERP non è attivato o l'utente non è amministratore
       tlbPrinc.FindControl(, idxBottoneDesigner).Visible = False
       tlbPrinc.FindControl(, idxBottoneUtenteModifica).BeginGroup = False 'Barra separatrice
    End If
    
    tlbQuality.Visible = False
    tlbDesigner.Visible = False
    CommandBars.TabWorkspace.ThemedBackColor = False

End Sub

Public Function FindCommandBar(ByVal ID As Long) As XtremeCommandBars.CommandBar
    Dim i As Long
    For i = 1 To CommandBars.Count
        If CommandBars(i).BarID = ID Then
            Set FindCommandBar = CommandBars(i)
            Exit For
        End If
    Next i
End Function


Public Property Get FunzioniM98() As CFunzioniMetodo98
    Set FunzioniM98 = mFunzioniM98
End Property

Public Property Get FormAttiva() As Object
    Set FormAttiva = MFormAttiva
End Property

Public Property Set FormAttiva(FRM As Object)
    Set MFormAttiva = FRM
End Property

Private Sub GestioneCmbDesignType(ControlCombo As XtremeCommandBars.CommandBarComboBox)
End Sub

Private Sub GestioneToolButDesigner(ByVal btnId As Long)
End Sub

Private Sub GestioneToolButQuality(ByVal btnId As Long)
End Sub

Private Sub MostraPopUpLingue()
    Dim PopupBar As CommandBar
    Dim Btn As CommandBarControl
    Dim strFile As String
    Dim strLingua As String
    Dim contIDLingua As Integer
    Dim PictLingua As StdPicture
    Dim objFS As New Scripting.FileSystemObject
    
    
    Set PopupBar = CommandBars.Add("PopupLingue", xtpBarPopup)
    With PopupBar.Controls
        contIDLingua = ID_POPUPLINGUEBASE
        strFile = Dir(MXNU.PercorsoFileLocali & "\MWRES??.DAT")
        While strFile <> ""
            strLingua = Mid(strFile, 6, 2)
            Set Btn = .Add(xtpControlButton, contIDLingua, strLingua)
            If objFS.FileExists(MXNU.PercorsoPgm & "\LanguageResources\Flags\Flag" & strLingua & ".ico") Then
            'If Dir$(MXNU.PercorsoPgm & "\LanguageResources\Flags\Flag" & strLingua & ".ico", vbNormal) <> "" Then
                Set PictLingua = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\Flag" & strLingua & ".ico")
                CommandBars.AddIconHandle PictLingua.Handle, contIDLingua, 0, False
            Else
                Set PictLingua = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\FlagEuro.ico")
                CommandBars.AddIconHandle PictLingua.Handle, contIDLingua, 0, False
            End If
            Btn.IconId = contIDLingua
            If MXNU.LinguaAttiva = strLingua Then Btn.Checked = True
            
            strFile = Dir()
            contIDLingua = contIDLingua + 1
        Wend
    End With
    Set objFS = Nothing
    MintIDLinguaMax = contIDLingua - 1
    PopupBar.ShowPopup
    
End Sub

Friend Sub SalvaLayout()
    Dim strCmdBarsState As String, strDockPaneState As String, strTemaAttivo As String, strColorScheme As String
    Dim f As Integer
    
    If GBolRipristinaLayout Then
        On Local Error Resume Next
        Kill MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini"
        On Local Error GoTo 0
        GBolRipristinaLayout = False
        Exit Sub
    End If
    
    strCmdBarsState = CommandBars.SaveStateToString(True, False, True)
    strDockPaneState = DockingPaneManager.SaveStateToString
    
    
    'Select Case CommandBars.VisualTheme
    Select Case LCase(GTemaAttivo)
        Case "office2007": strTemaAttivo = "Office2007"
        Case "winxp.luna": strTemaAttivo = "WinXPLuna"
        Case "winxp.royale": strTemaAttivo = "WinXPRoyale"
        Case "vista": strTemaAttivo = "Vista"
        Case "": strTemaAttivo = "system"
    End Select
        
'    On Local Error Resume Next
'    If Dir$(MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini", vbNormal) <> "" Then
'        Kill MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini"
'    End If
'    On Local Error GoTo 0
    
    f = FreeFile
    Open MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini" For Output As #f
    Print #f, "CommandBars=" & strCmdBarsState
    Print #f, "DockingPane=" & strDockPaneState
    Print #f, "TemaAttivo=" & strTemaAttivo
    If DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden Then    'Anomalia 10154
        DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden = False
        DoEvents
        Print #f, "SplitterPreferiti=" & frmModuli.MWSplitter1.SplitterTop + 2   '<<< Salvo con 2 pixel di "lasco" altrimenti poi non si riposiziona correttamente
        DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hide
    Else
        Print #f, "SplitterPreferiti=" & frmModuli.MWSplitter1.SplitterTop + 2   '<<< Salvo con 2 pixel di "lasco" altrimenti poi non si riposiziona correttamente
    End If
    Print #f, "IconeGrandi=" & Abs(frmModuli.IconeGrandi)
    Print #f, "AnimazionePannelli=" & Abs(metodo.DockingPaneManager.Options.AnimationDelay)
    Print #f, "PannelliVisibili=" & frmModuli2005.ShortcutBar1.ExpandedLinesCount
    Print #f, "UsaResizer=" & Abs(MXNU.UsaResizer)
'    Select Case CommandBarsGlobalSettings.ColorManager.SystemTheme
'        Case xtpSystemThemeAuto, xtpSystemThemeUnknown: strColorScheme = "Default"
'        Case xtpSystemThemeBlue: strColorScheme = "Blue"
'        Case xtpSystemThemeSilver: strColorScheme = "Silver"
'        Case xtpSystemThemeOlive: strColorScheme = "Olive"
'        Case Else: strColorScheme = "Default"
'    End Select
'    Print #f, "SchemaColori=" & strColorScheme
    
    Close #f

End Sub

'Salva posizione, dimensione e stato di Toolbar, MenuBar e DockingPane
Friend Sub SalvaLayout_OLD()
    Dim strCmdBarsState As String, strDockPaneState As String, strTemaAttivo As String, strColorScheme As String
    Dim f As Integer
    
    If GBolRipristinaLayout Then
        On Local Error Resume Next
        Kill MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini"
        On Local Error GoTo 0
        GBolRipristinaLayout = False
        Exit Sub
    End If
    
    strCmdBarsState = CommandBars.SaveStateToString(True, False, True)
    strDockPaneState = DockingPaneManager.SaveStateToString
    
    Select Case CommandBars.VisualTheme
        Case xtpThemeOffice2000: strTemaAttivo = "Office2000"
        Case xtpThemeOfficeXP: strTemaAttivo = "OfficeXP"
        Case xtpThemeOffice2003: strTemaAttivo = "Office2003"
        Case xtpThemeNativeWinXP: strTemaAttivo = "XPNative"
    End Select
        
'    On Local Error Resume Next
'    If Dir$(MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini", vbNormal) <> "" Then
'        Kill MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini"
'    End If
'    On Local Error GoTo 0
    
    f = FreeFile
    Open MXNU.PercorsoPreferenze & "\Layout_" & MXNU.UtenteAttivo & ".ini" For Output As #f
    Print #f, "CommandBars=" & strCmdBarsState
    Print #f, "DockingPane=" & strDockPaneState
    Print #f, "TemaAttivo=" & strTemaAttivo
    Print #f, "SplitterPreferiti=" & frmModuli.MWSplitter1.SplitterTop + 2   '<<< Salvo con 2 pixel di "lasco" altrimenti poi non si riposiziona correttamente
    Print #f, "IconeGrandi=" & Abs(frmModuli.IconeGrandi)
    Print #f, "AnimazionePannelli=" & Abs(metodo.DockingPaneManager.Options.AnimationDelay)
    Print #f, "PannelliVisibili=" & frmModuli2005.ShortcutBar1.ExpandedLinesCount
    Print #f, "UsaResizer=" & Abs(MXNU.UsaResizer)
    Select Case CommandBarsGlobalSettings.ColorManager.SystemTheme
        Case xtpSystemThemeAuto, xtpSystemThemeUnknown: strColorScheme = "Default"
        Case xtpSystemThemeBlue: strColorScheme = "Blue"
        Case xtpSystemThemeSilver: strColorScheme = "Silver"
        Case xtpSystemThemeOlive: strColorScheme = "Olive"
        Case Else: strColorScheme = "Default"
    End Select
    Print #f, "SchemaColori=" & strColorScheme
    
    Close #f
End Sub

Public Property Get UnloadingMetodo() As Boolean
    UnloadingMetodo = mBolUnloadingMetodo
End Property

Public Property Get LoadingMetodo() As Boolean
    LoadingMetodo = mBolLoadingMetodo
End Property

Public Property Get pBottoniAgenti() As setBottoneAgente
    pBottoniAgenti = mEnmBottoneAgenti
End Property

Public Property Let pBottoniAgenti(new_Valore As setBottoneAgente)
    mEnmBottoneAgenti = new_Valore
End Property

Public Property Get FormsCount() As Integer
    FormsCount = Forms.Count
End Property
Public Property Let UnloadingMetodo(new_UnloadingMetodo As Boolean)
    mBolUnloadingMetodo = new_UnloadingMetodo
End Property



'Private Sub Barra_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If (Button = vbRightButton) And (Barra.Buttons(idxBottoneDefAgenti).Value = tbrUnpressed) And _
'        (x >= Barra.Buttons(idxBottoneDefAgenti).Left And x <= (Barra.Buttons(idxBottoneDefAgenti).Left + Barra.Buttons(idxBottoneDefAgenti).Width)) Then
'        If (mEnmBottoneAgenti = ageCollegamenti) Then
'            Barra.Buttons(idxBottoneDefAgenti).Image = 16
'            mEnmBottoneAgenti = ageMostraRiferimenti
'        ElseIf (mEnmBottoneAgenti = ageMostraRiferimenti) Then
'            Barra.Buttons(idxBottoneDefAgenti).Image = 18
'            mEnmBottoneAgenti = ageMostraCampiDB
'        ElseIf (mEnmBottoneAgenti = ageMostraCampiDB) Then
'            Barra.Buttons(idxBottoneDefAgenti).Image = 21
'            mEnmBottoneAgenti = ageDipendenze
'        Else
'            Barra.Buttons(idxBottoneDefAgenti).Image = 15
'            mEnmBottoneAgenti = ageCollegamenti
'        End If
'    End If
'End Sub




'Utilizzata all'interno di MXDesigner
Public Sub VisualizzaCoordDesigner(ByVal TipoCoord As String, Valore As Variant)
    Select Case UCase(TipoCoord)
        Case "TOP": FindCommandBar(ID_TLB_DESIGNER).FindControl(, idxDesignerLabelTop).Caption = "Top: " & Valore
        Case "LEFT": FindCommandBar(ID_TLB_DESIGNER).FindControl(, idxDesignerLabelLeft).Caption = "Left: " & Valore
        Case "HEIGHT": FindCommandBar(ID_TLB_DESIGNER).FindControl(, idxDesignerLabelHeight).Caption = "Height: " & Valore
        Case "WIDTH": FindCommandBar(ID_TLB_DESIGNER).FindControl(, idxDesignerLabelWidth).Caption = "Width: " & Valore
    End Select
End Sub

Private Sub BarraStato_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

#If TOOLS = 0 Then
        If x >= BarraStato.Panels("Lingua").Left And x <= BarraStato.Panels("Lingua").Left + BarraStato.Panels("Lingua").Width And MXNU.MetodoXP Then
            Call frmpopup.MostraPopUpLingue(Me, "Lingue")
            'Call MDIForm_Resize
        
    #If ISMETODOXP = 1 Then
        ElseIf x >= BarraStato.Panels("Designer").Left And x <= BarraStato.Panels("Designer").Left + BarraStato.Panels("Designer").Width And MXNU.MetodoXP Then
            If ((MXNU.ControlloModuliChiave(modMyErpRunTime) = 0) _
                Or (MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0) _
                Or (MXNU.ControlloModuliChiave(modMyErpDesign) = 0)) Then
                
                Call frmpopup.MostraPopUpDesigner(Me, "Designer")
                'Call MDIForm_Resize
            End If
    #End If
        
        End If
    
#End If
    
End Sub


'DESIGNER:  A seconda della scelta del valore del combo l'utente va in Design dei controlli
'           dell'oggetto contentitore del controllo attivo oppure delle linguette in testa all'oggetto contenitore
'Private Sub cmbDesignType_Click()
'  Dim objManagedControl As Object
'  Dim oCtrl As Object
'  Dim i As Integer
'  Dim bolLing As Boolean
'
'  If Not (FormAttiva Is Nothing) Then
'      Select Case cmbDesignType.ListIndex
'            Case 0      ' CONTROLLI
'                metodo!Barra.Buttons(idxBottoneDesigner).Enabled = False
'                With tlbDesigner.Buttons
'                    For i = 1 To .Count
'                        .Item(i).Enabled = True
'                    Next i
'                    'Disattivo il bottone showdiff
'                    .Item("showdiff").Enabled = False
'                    .Item("setting").Enabled = False
'                    .Item("unvisible").Value = tbrUnpressed
'                    .Item("disable").Value = tbrUnpressed
'                    .Item("showdiff").Value = tbrUnpressed
'                End With
'
'                ' Vado sui controlli
'                If TypeName(FormAttiva.ActiveControl) = "MWLinguetta" Then
'                    ' Se sono su linguetta cerco la scheda con lo stesso indice...
'                    For Each oCtrl In FormAttiva.Controls
'                        If TypeName(oCtrl) = "MWSchedaBox" And HasIndex(oCtrl) Then
'                            If oCtrl.Index = FormAttiva.ActiveControl.Index Then
'                                Set objManagedControl = oCtrl
'                                Exit For
'                            End If
'                        End If
'                    Next
'                ElseIf TypeName(FormAttiva.ActiveControl) = "MWSchedaBox" Then
'                    ' Se il focus è su scheda allora in container è la scheda stessa
'                    Set objManagedControl = FormAttiva.ActiveControl
'                Else
'                    ' Altrimenti prendo il container del controllo attivo
'                    Set objManagedControl = FormAttiva.ActiveControl.Container
'                End If
'                bolLing = False
'
'            Case 1      ' LINGUETTE
'                metodo!Barra.Buttons(idxBottoneDesigner).Enabled = False
'
'                With tlbDesigner.Buttons
'                    For i = 1 To .Count
'                        .Item(i).Enabled = True
'                    Next i
'                    'Disattivo il bottone showdiff
'                    .Item("showdiff").Enabled = False
'                    .Item("setting").Enabled = False
'                    .Item("unvisible").Value = tbrUnpressed
'                    .Item("disable").Value = tbrUnpressed
'                    .Item("showdiff").Value = tbrUnpressed
'
'                End With
'
'                ' Vado sulle linguette (devo iterare fino a trovare ....
'                ' ... un contenitore delle linguette a partire dal controllo attivo)
'                Dim objChecked As Object
'                Dim bolFind As Boolean
'
'                bolFind = False
'                Set objChecked = MFormAttiva.ActiveControl
'                Do While Not (bolFind)
'                    If TypeName(objChecked) = "MWLinguetta" Then
'                        ' Se ho il focus sulla linguetta vado in design sul suo contenitore
'                        Set objManagedControl = objChecked.Container
'                        bolFind = True
'                    Else
'                        If TypeName(objChecked) <> "MWSchedaBox" And TypeName(objChecked) <> "Frame" And objChecked.hwnd <> MFormAttiva.hwnd Then
'                            ' Se il controllo non è un contenitore mi posiziono sul container
'                            Set objChecked = objChecked.Container
'                        End If
'                        If objChecked.hwnd = MFormAttiva.hwnd Then
'                            ' Sono arrivato al contenitore "form"
'                            Set objManagedControl = MFormAttiva
'                            bolFind = True
'                        Else
'                            'Nel caso di oggetto frame (non supporta la collection controls) .....
'                            If TypeName(objChecked) = "Frame" Then
'                                For Each oCtrl In objChecked.Container.Controls
'                                    If TypeName(oCtrl) = "MWLinguetta" And oCtrl.Container.Name = objChecked.Name Then
'                                        Set objManagedControl = objChecked
'                                        bolFind = True
'                                        Exit For
'                                    End If
'                                Next
'                            Else
'                                ' Sono nella scheda ...
'                                For Each oCtrl In objChecked.Controls
'                                    If TypeName(oCtrl) = "MWLinguetta" Then
'                                        ' Il contenitore contiene almeno una linguetta
'                                        Set objManagedControl = objChecked
'                                        bolFind = True
'                                        Exit For
'                                    End If
'                                Next
'                            End If
'                        End If
'                    End If
'                    If Not bolFind Then
'                        Set objChecked = objChecked.Container
'                    End If
'                Loop
'                Set objChecked = Nothing
'                bolLing = False
'
'        End Select
'
'        cmbDesignType.Enabled = False
'        If cmbDesignType.ListIndex <> -1 Then
'            ' Visualizzo la scheda contenente le coordinate
'            SchedaDesigner(1).Visible = True
'            If Not Designer.InitDesigner(objManagedControl, bolLing) Then
'                Call MXNU.MsgBoxEX(2630, vbExclamation, 1007)
'            End If
'        Else
'            ' Nascondo la scheda contenente le coordinate
'            SchedaDesigner(1).Visible = False
'        End If
'
'        Set objManagedControl = Nothing
'
'  End If
'
'End Sub
'
Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    'Si accede agli elementi di una XtremeCommandBar solo per Id, quindi in base al range dell'id so su che toolbar o menu è stata effettuata la scelta
    Select Case Control.ID
        'Case idxBottoneModuli To idxBottoneDesigner, idxBottoneUtenteModifica To idxBottoneSchedula, idxBottoneAttivaDesigner To idxBottoneGrigliaDesigner, idxBottoneDefAgenti, idxBottoneNomiCtrlCmp To idxBottoneRicProfili, idxBottoneZoom100 To idxBottoneZoom200, ID_TLBITEM_INFO, ID_TLBITEM_PREFERENCES, ID_TLBITEM_CHANGETHEME To ID_TLBITEM_THEME_NATIVEXP, ID_TLBITEM_FRMORIGINALSIZE, ID_TLBITEM_THEME_OFFICE2007
        Case idxBottoneModuli To idxBottoneDesigner, idxBottoneUtenteModifica To idxBottoneSchedula, idxBottoneAttivaDesigner To idxBottoneGrigliaDesigner, idxBottoneDefAgenti, idxBottoneNomiCtrlCmp To idxBottoneRicProfili, idxBottoneZoom100 To idxBottoneZoom200, ID_TLBITEM_INFO, ID_TLBITEM_PREFERENCES, ID_TLBITEM_CHANGETHEME To ID_TLBITEM_THEME_VISTA, ID_TLBITEM_FRMORIGINALSIZE, ID_TLBITEM_THEME_SYSTEM, ID_TLBITEM_THEME_OFFICE2010
            Call GestioneToolBut2005(Control.ID)
        Case idxQualityNew To idxQualityPrintList
            Call GestioneToolButQuality(Control.ID)
        Case idxDesignerCmbDesignType To idxDesignerSaveVersion
            Call GestioneToolButDesigner(Control.ID)
        Case ID_POPUPLINGUEBASE To MintIDLinguaMax
            Call CambiaLingua(Control.Caption)
        Case Else
           'MenuBar
            Dim objMenu As CGestMenu
            For Each objMenu In McolMenu
                If objMenu.MenuID = Control.ID Then
                    AttivaVoceMenu objMenu.key
                    Exit For
                End If
            Next
    End Select
    
End Sub

Private Sub CommandBars_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    'If BarraStato.Visible Then
    '    Bottom = BarraStato.Height
    'End If

End Sub

Private Sub CommandBars_SpecialColorChanged()
    'Call CambiaSchemaColori((CommandBars.VisualTheme = xtpThemeOffice2003), metodo.CommandBars.GetSpecialColor(XPCOLOR_TOOLBAR_FACE), metodo.CommandBars.GetSpecialColor(XPCOLOR_SPLITTER_FACE))
    Call CambiaColoriFormAttive
    Call InizializzaSpread(True, True, SysGradientColor1)
    'Inizializzo lo spread anche per MXKit e per MXBusiness
    Call MXVI.InitSpreadEvolus
    Call MXBusiness.InitSpreadEvolus
    
End Sub

Private Sub DockingPaneManager_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    If Action = PaneActionClosed Then
        If frmLog.IsLogPane(Pane.ID) Then
            Call frmLog.ScaricaLog(Pane.ID)
            metodo.DockingPaneManager.DestroyPane Pane
        End If
    ElseIf Action = PaneActionPinned Then
        If Pane.ID = ID_PANE_NAVBAR Then
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneModuli).Checked = False
        End If
    ElseIf Action = PaneActionDocked Then
        If Pane.ID = ID_PANE_NAVBAR Then
            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneModuli).Checked = True
            If frmModuli2005.ShortcutBar1.Selected.ID = ID_BAR_PROGMODULES Then
                Call DaiFocusAlberoModuli
            End If
        End If
    End If

End Sub

Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = ID_PANE_NAVBAR Then
        Item.Handle = frmModuli2005.hwnd
    ElseIf (Item.ID = ID_PANE_TASKBAR2) Then
        If (Not mMessagingEngine Is Nothing) Then
            Item.Handle = mMessagingEngine.GetMailingFormHandle()
        End If
    End If
End Sub

Private Sub MDIForm_Initialize()
    InitCommonControls
End Sub

Private Sub MDIForm_Load()
Dim intq As Integer
    
    mBolLoadingMetodo = True
    Me.Hide
    Set MXNU.FrmMetodo = Me
    UnloadingMetodo = False
    
    metodo.MousePointer = vbHourglass
    
    '>>>ALLINEO I MENU A DESTRA
    'mnuAiuto.Caption = Chr$(8) & "&?  "
    
    '>>>allineo la barra degli strumenti
    'Barra.Align = Val(MXNU.LeggiProfilo(MXNU.File_ini_personale$, "barrasts", "align", "1"))
    'Barra.Buttons(idxBottoneSpostaBarra).Image = Barra.Align + 10
    
  
#If TOOLS <> 1 Then
    '******************************************************
    'MODIFICHE PER EVITARE IL CONTROLLO DELLA CHIAVE IN TSE
    Dim strResult  As String
    strResult = MXNU.LeggiProfilo(MXNU.DirAvvio & "\mw.ini", "METODOW", "CONTROLLOCHIAVE", "TRUE")
    mBolFaiControlloChiaveTSE = True
    If strResult = "FALSE" Then
        mBolFaiControlloChiaveTSE = False
    End If
    '******************************************************
    'If MXNU.MetodoXP Then
    '    Me.Icon = Me.ImglistBottoniXP.ListImages("IconaMXP").Picture
    'End If
#Else
    Me.Icon = Me.ImmaginiBottoni.ListImages("TOOLS").Picture
#End If

    Set Barra = New cBarra
    
    Call CreateToolBars
    Call CreatePanes
       
    '>>>barra di stato
    Set BarraStato = New CBarraStato
    Call CreateStatusBar
    Call AggiornaStatusBar
    
    '>>>carico l'albero dei moduli
    Call CaricaModuliMetodo
    
    'Unload frmIntro
    frmIntro.MostraMessaggioOperazione (12078)
    DoEvents
    '>>>mostro la form
    mBolLoadingMetodo = False
    'Call frmModuli.Ridimensiona
        
    'controllo modulo jobschedulatr, se non c'è disabilito il botone della toolbar
    'metodo.Barra.Buttons(idxBottoneSchedula).Enabled = (MXNU.ControlloModulichiave(modJobScheduler) = 0)
    Call CaricaLayout
    
    On Local Error Resume Next
    Dim ScreenSizeX As Long, ScreenSizeY As Long
    ScreenSizeX = (Screen.Width \ Screen.TwipsPerPixelX)
    ScreenSizeY = (Screen.Height \ Screen.TwipsPerPixelY)
        
    '===========================================================================================================
    'SFONDO APPLICAZIONE
    '===========================================================================================================
    'sfondo personalizzato per versione
    Me.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondoEvolus." & MXNU.VersioneAttiva & ".jpg")
    If (Err <> 0) Then
        Err.Clear
        Set Me.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondoEvolus." & MXNU.VersioneAttiva & ".bmp")
    End If
    'sfondo personalizzato per utente
    If (Err <> 0) Then
        Err.Clear
        Me.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondoEvolus." & MXNU.UtenteAttivo & ".jpg")
    End If
    If (Err <> 0) Then
        Err.Clear
        Set Me.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondoEvolus." & MXNU.UtenteAttivo & ".bmp")
    End If
    'sfondo standard per risoluzione video
    If (Err <> 0) Then
        Err.Clear
        'Me.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo2005" & ScreenSizeX & "x" & ScreenSizeY & ".jpg")
        Me.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\SfondoEvolus" & ScreenSizeX & "x" & ScreenSizeY & ".jpg")
    End If
    If (Err <> 0) Then
        Err.Clear
        Set Me.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondoEvolus.jpg")
    End If
    If (Err <> 0) Then
        Err.Clear
        Set Me.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondoEvolus.bmp")
    End If
    '===========================================================================================================
    
    Me.Show
    CommandBars.TabWorkspace.Refresh   'Altrimenti non si disegna correttamente l'immagine di sfondo della MDI
    DoEvents
    Unload frmIntro
    
    'Se l'utente ha chiuso il pannello dei moduli lo nascondo solo dopo aver visualizzato la form altrimenti la barra di stato si visualizza a metà videata (!?!)
    If Not MDockAperto Then DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hide
    On Local Error GoTo 0
    DoEvents
    Call MXNU.Attiva_Toolbar(hwnd, 0)
    DoEvents
    'attivo l'applicazione
    On Local Error Resume Next
    If Me.Enabled Then Call AppActivate(Me.Caption)
    On Local Error GoTo 0
    
    Set mFunzioniM98 = New CFunzioniMetodo98
    metodo.MousePointer = vbDefault
    DoEvents
    
    ' rif.svil. S1196
    Call ControllaOpSchedulate
    
    'Call CaricaLayout
    DoEvents
    On Local Error Resume Next
    'frmModuli.ModuloAttivo = "Comuni"
    Set MEmptyform = New EmptyForm
    Load MEmptyform
    'Aggiungo la form Empty alla collection delle form da non scaricare mai
    McolFormsInNavBar.Add MEmptyform.Name, MEmptyform.Name
    
    'Anomalia 8516 - Spostato nel main, dopo il load della MDI
    'frmModuli.ModuloAttivo = "*DA_INI*"    'Leggo il modulo attivo dall'mw.ini. Vedi funzione AttivaMenu su MetodoXP
    On Local Error GoTo 0
            
    '>>> sistema di messaggistica
    If (Not mMessagingEngine Is Nothing) Then
        Set mMessagingEngine.mMDiForm = Me
        If (mMessagingEngine.IsMessagingSystemInstalled()) Then
            'apro le note nel desktop
            Call mMessagingEngine.ShowStartupNotes
        End If
    End If
    
    If MLngSplitter <> 0 Then
        If DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden Then   'Anomalia 10154
            DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden = False
            DoEvents
            frmModuli.MWSplitter1.SplitterTop = MLngSplitter
            DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hide
        Else
            frmModuli.MWSplitter1.SplitterTop = MLngSplitter
        End If
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim r%
If Not GBolNoMsgConfermaUscita Then
        'If InStampa Then
        If MXCREP.StampaInCorso Then
           Call MXNU.MsgBoxEX(1096, vbOKOnly + vbExclamation, 1007)
           Cancel = True
        Else
            r = MXNU.MsgBoxEX(1645, 36, 1007, Array(Me.Caption))
            If r = IDNO Then
                Cancel = True
            End If
        End If
End If
End Sub

'Private Sub MDIForm_Resize()
'
'    On Local Error Resume Next
'
'    #If ISMETODO2005 <> 1 Then
'        If Not frmModuli Is Nothing Then
'            #If ISMETODOXP = 1 Then
'                If MXNU.MetodoXP Then
'                    frmModuli.Height = Me.ScaleHeight
'                    'frmToolTip.Height = Me.ScaleHeight
'                Else
'                    frmModuli.Ridimensiona
'                End If
'            #Else
'                frmModuli.Ridimensiona
'            #End If
'        End If
'    #End If
'    On Local Error GoTo 0
'
'End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    Set Barra = Nothing
    UnloadingMetodo = True
    Set frmLog = Nothing    '<<< Da richiamare PRIMA di SalvaLayout, altrimenti vengono salvati anche i panes dei log aperti
    'Da richiamare prima di SalvaLayout altrimenti viene salvato anche il pane del browser
    If (Not mMetodoBrowser Is Nothing) Then
        mMetodoBrowser.Dispose
        Set mMetodoBrowser = Nothing
    End If
    Call SalvaLayout
    If (Not frmModuli Is Nothing) Then
        Unload frmModuli
    End If
    DoEvents
    Set mFunzioniM98 = Nothing
    Set MXNU.FrmMetodo = Nothing
    Set frmIntro = Nothing
    MbolInChiusura = True
    Call ChiudiDitta
    Call ChiudiMetodo
    Set McolFormsInNavBar = Nothing
    Set metodo = Nothing

End Sub


Private Sub MFormAttiva_Deactivate()
    ' Riga commentata per DESIGNER; ATTENZIONE : MXNU.ImpostaFormAttiva non imposta se c'è questo SET !!!
    ' Chi de-commenta la riga dovrà passare sul mio cadavere ---- DANIEL
    ' Set MFormAttiva = metodo.ActiveForm
End Sub

Private Sub MFormAttiva_Unload(Cancel As Integer)
'    #If ISMETODO2005 = 1 Then
'        On Local Error Resume Next
'        If (Not MFormAttiva.mResize Is Nothing) Then
'            MFormAttiva.mResize.Terminate
'            Set MFormAttiva.mResize = Nothing
'        End If
'        On Local Error Resume Next
'    #End If
 
End Sub

Private Sub mnuAiutoItem_Click(Index As Integer)
    'Call EseguiAzione_I("AiutoItem", Index, mnuAiutoItem(Index).HelpContextID)
End Sub

'Private Sub mnuMenuItem_Click(Index As Integer)
'     Select Case Index
'         Case 0
'             If ApriDittaAnno(False, "", "") Then
'                Call MDIForm_Resize
'             End If
'         Case 1
'             Dim NuovoAnno As Integer
'             If SelezioneAnno(True, NuovoAnno) Then
'                 Call ChiudiFormAttive
'                 MXNU.AnnoAttivo = NuovoAnno   'Rif. Sk. Anomalie Nr. 5113
'                 Call ApriAnno(True)
'             End If
'         Case 3
'            #If ISMETODOXP = 1 Then
'                If MXNU.MetodoXP Then
'                    If frmModuli.StatoMenuAlbero(0) = False Then
'                        'il menu moduli non e' aperto ne lockato
'                        frmModuli.StatoMenuModuli = MnuModATTIVO
'                        Call frmModuli.AlberoBloccatoMenu(True, 0)
'                    Else
'                        'il menu moduli e' gia' aperto e lockato
'                        'frmModuli.StatoMenuModuli = MnuModNONATTIVO
'                        Call frmModuli.AlberoBloccatoMenu(False, 0)
'                    End If
'                Else
'                    SetCursorPos 10, ((metodo.Height - metodo.ScaleHeight) + 30) / Screen.TwipsPerPixelY
'                    Call AttivaMenuMetodo
'                End If
'
'            #Else
'                 SetCursorPos 10, ((metodo.Height - metodo.ScaleHeight) + 30) / Screen.TwipsPerPixelY
'                 Call AttivaMenuMetodo
'
'            #End If
'
'         Case 5
'             Call CambioUtenteAttivo
'             Call MDIForm_Resize
'         Case 7 'Uscita
'             Unload Me
'    End Select
'
'End Sub
'
Public Function EseguiAzione(ByVal NomeMenu As Variant, ByVal indice As Integer, HelpContextID As Long)
    
    On Local Error Resume Next
    frmModuli.Enabled = False
    On Local Error GoTo 0
    
    On Local Error GoTo err_EseguiAzione
    
On Local Error Resume Next

    Dim k$
    k = McolMenu(NomeMenu & "_" & indice).key
    If Err Then  'Voce di menu nascosta
        Call EseguiAzione_i(NomeMenu, indice, HelpContextID)
    Else
        On Local Error GoTo err_EseguiAzione
        If frmModuli.TrwModuli.Nodes(NomeMenu & "_" & indice).key <> "" Then
            Call EseguiAzione_i(NomeMenu, indice, HelpContextID)
        End If
    End If
    
err_EseguiAzione:
    If Err <> 0 Then Call MXNU.MsgBoxEX(1595, vbInformation + vbOKOnly, 1007)
    
    On Local Error Resume Next
    If Not UnloadingMetodo Then
        frmModuli.Enabled = True
    End If
    On Local Error GoTo 0

End Function

Private Sub ChiudiEsercizioTransitorio()
'    Dim hdb%, r%
'
'    If MsgBox("CHIUDERE TUTTI GLI ESEGUIBILI DI METODO!!!" & Chr$(10) & "Confermi la chiusura dell'anno transitorio?", MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON2, "ATTENZIONE") = IDNO Then
'       Exit Sub
'    End If
'
'    hdb = MXDB.dbConnetti(MXDB.hAmbiente, metodo.hwnd, "DSN=" & MXNU.DittaAttiva & CStr(Val(MXNU.AnnoAttivo) - 1), 60, 0, 0, 1, vbUseServersideCursor)
'    If hdb Then
'        r = MXDB.dbEseguiSQL(hdb, "UPDATE TabDitte SET StatoAnno=3")
'        r = MXDB.dbEseguiSQL(hndDBArchivi, "UPDATE TabDitte SET StatoAnno=1")
'        r = MXDB.dbChiudiDB(hdb)
'        metodo.mnuDitteItem(ID_MNUDITTEITEM_CHIUDITRANS).Visible = False
'        metodo.mnuDitteItem(ID_MNUDITTEITEM_CHIUDITRANS - 1).Visible = False
'        MXNU.Stato_Esercizio = 1
'        metodo.lbDittaAnno.Caption = DammiStatusBar()
'    Else
'        MsgBox "Impossibile stabilire la connessione con l'esercizio precedente; operazione fallita!", MB_ICONSTOP, "Chiusura Esercizio Transitorio"
'    End If
'
End Sub

Private Sub MFormAttiva_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    Dim frmAt As Form
'    Dim bolTrovato As Boolean
'    For Each frmAt In VB.Forms
'        If frmAt.Name = MFormAttiva.Name Then
'            bolTrovato = True
'            Exit For
'        End If
'    Next frmAt
    
'    If Not bolTrovato Then Exit Sub
    
    If Shift = vbCtrlMask Then
        Dim bolEseguiKD As Boolean
        On Local Error Resume Next
        bolEseguiKD = True
        'bolEseguiKD = (Not ActiveForm.MWAgt1.NOFormKeyDown(ActiveForm.hwnd, KeyCode, Shift))
        'If (bolEseguiKD) Then
            Select Case KeyCode
                Case BTN_PRIMO
                    GestioneToolBut2005 idxBottonePrimo
                    KeyCode = 0
                Case BTN_ULTIMO
                    GestioneToolBut2005 idxBottoneUltimo
                    KeyCode = 0
                Case BTN_PREC
                    GestioneToolBut2005 idxBottonePrecedente
                    KeyCode = 0
                Case BTN_SUCC
                    GestioneToolBut2005 idxBottoneSuccessivo
                    KeyCode = 0
                Case BTN_ANN
                    GestioneToolBut2005 idxBottoneAnnulla
                    KeyCode = 0
                Case BTN_REG
                    GestioneToolBut2005 idxBottoneRegistra
                    KeyCode = 0
                Case BTN_INS
                    GestioneToolBut2005 idxBottoneInserisci
                    KeyCode = 0
                Case BTN_MOD
                    GestioneToolBut2005 idxBottoneDettaglio
                    KeyCode = 0
                Case BTN_VISUTMOD
                    GestioneToolBut2005 idxBottoneUtenteModifica
                    KeyCode = 0
                Case BTN_ALLINONE
                     If ((MXNU.ControlloModuliChiave(modAllInOneRuntime) = 0) _
                        Or (MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0)) Then
                        
                        GestioneToolBut2005 idxBottoneAllInOne
                        KeyCode = 0
                    End If
                Case BTN_ZOOM
                    'Scheda sviluppo n.ro 1128 - Gestione zoom
                    'If MXNU.ControlloModulichiave(modMyErpRunTime) = 0 Then
                        GestioneToolBut2005 idxBottoneZoom
                        KeyCode = 0
                    'End If
                Case BTN_TROVA
                    GestioneToolBut2005 idxBottoneTrova
                    KeyCode = 0
                Case BTN_DESIGNER
                    'DESIGNER
                    If MXNU.ControlloModuliChiave(modMyErpDesign) = 0 Then
                        GestioneToolBut2005 idxBottoneDesigner
                        KeyCode = 0
                    End If
            End Select
        'End If
    ElseIf Shift = vbCtrlMask + vbShiftMask Then
        Select Case KeyCode
            Case vbKeyC
                '***Barra.Buttons(idxBottoneDefAgenti).Value = tbrUnpressed
                'Call Form.AzioniMetodo(MetFMostraCampiDBAnagr)
                If Not MXNU.CtrlAccessi Then    ' Rif. anomalia n.ro 4824
                    'GestioneToolBut2005 Barra.Buttons(idxBottoneDefAgenti).ButtonMenus(3)
                    GestioneToolBut2005 idxBottoneSituazAnagr
                End If
                KeyCode = 0
        End Select
    Else
        Select Case KeyCode
            Case vbKeyF1
                'Call ApriHelp(False)
                GestioneToolBut2005 idxBottoneHelp
                KeyCode = 0
        End Select
    End If

End Sub

Private Sub MFormAttiva_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If (Me.Barra.Buttons(idxBottoneDefAgenti).Value = 1) Then
'        Me.Barra.Buttons(idxBottoneDefAgenti).Value = 0
'        Me.pBottoniAgenti = Me.pBottoniAgenti - 1
'    End If
    'If FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneAttivaDefAgenti).DefaultItem Then
    'If FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneDefAgenti).CommandBar.FindControl(, idxBottoneAttivaDefAgenti).DefaultItem Then
    '    FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneDefAgenti).CommandBar.FindControl(, idxBottoneAttivaDefAgenti).DefaultItem = False
    '    Me.pBottoniAgenti = Me.pBottoniAgenti - 1
    'End If
End Sub

'Private Sub MFormAttiva_Unload(Cancel As Integer)
'    Dim x
'    x = x
    'Call MXNU.Attiva_Toolbar(0, 0)
'    Set MFormAttiva = Me.ActiveForm
'End Sub

'Public Function MDICaricaTabella(ByVal strNomeDbTabella As String, ByVal vntDesTabella As Variant, ByVal lngHelp As Long) As Boolean
'
'    On Local Error Resume Next
'    Call CaricaTabella(strNomeDbTabella, vntDesTabella, lngHelp)
'    MDICaricaTabella = (Err = 0)
'    On Local Error GoTo 0
'
'End Function
'
'Public Function MDIFiltro(ByVal NomeFiltro As String, ByVal HelpContextID As Long) As Boolean
'
'    On Local Error Resume Next
'    Call PrintBlaster(NomeFiltro, HelpContextID)
'    MDIFiltro = (Err = 0)
'    On Local Error GoTo 0
'
'End Function
'
Public Property Get MDIForms(Optional ByVal key As Variant) As Object

    Set MDIForms = Nothing
    On Local Error Resume Next
    If IsMissing(key) Then
        Set MDIForms = VB.Forms
    Else
        Set MDIForms = VB.Forms(key)
    End If

End Property

Public Function mdiEstensioneAgenti(pobjAgt As MXKit.CAgenteAuto, ByVal strComando As String, ByVal strParametri As String) As Boolean
    Dim vntParam As Variant
    Dim bolAspetta As Boolean
    Dim bolRes As Boolean
    Dim lngHwnd As Long
    Dim i&
    Dim strAzione As String
    Dim lngIDAzione As Long
    Dim lngHelpIDAzione As Long
            
    bolRes = True
    On Local Error GoTo err_mdiEstensioneAgenti
    Select Case strComando
        Case "#ESEGUIMENU"
            strAzione = i2s(strParametri, 1, ";")
            lngIDAzione = Val(i2s(strParametri, 2, ";"))
            lngHelpIDAzione = Val(i2s(strParametri, 3, ";"))
            If i2s(strParametri, 4, ";") <> "" Then bolAspetta = CBool(i2s(strParametri, 4, ";"))   ' Attesa opzionale
            On Local Error GoTo err_EseguiAzione
            Call MXNU.FrmMetodo.EseguiAzione(strAzione, lngIDAzione, lngHelpIDAzione)
            On Local Error GoTo err_mdiEstensioneAgenti
            If bolAspetta <> 0 Then
                
                With VB.Forms
                    For i = 1 To .Count
                        If (.Item(i).HelpContextID = lngHelpIDAzione) Then
                            lngHwnd = .Item(i).hwnd
                            Exit For
                        End If
                    Next i
                End With
                If lngHwnd <> 0 Then
                    Do While (IsWindow(lngHwnd))
                        DoEvents
                        ' Rif. scheda #9058 (uso insulso di cicli CPU)
                        Sleep 200
                    Loop
                End If
            End If
        
        Case "#TABELLA"
            vntParam = Split(strParametri, ";")
            If UBound(vntParam) < 3 Then
                ReDim Preserve vntParam(3)
                vntParam(3) = 0
            ElseIf UBound(vntParam) < 4 Then
                ReDim Preserve vntParam(4)
                vntParam(4) = ""
            End If
            bolAspetta = CBool(vntParam(3))
            Call CaricaTabella((vntParam(0)), (vntParam(1)), (vntParam(2)), , vntParam(4))
            If bolAspetta <> 0 Then
                metodo.MousePointer = vbDefault
                With VB.Forms
                    For i = 1 To .Count
                        If (.Item(i).HelpContextID = Val(vntParam(2))) Then
                            lngHwnd = .Item(i).hwnd
                            Exit For
                        End If
                    Next i
                End With
                If lngHwnd <> 0 Then
                    Do While (IsWindow(lngHwnd))
                        DoEvents
                        ' Rif. scheda #9058 (uso insulso di cicli CPU)
                        Sleep 200
                    Loop
                End If
            End If
        
        Case "#STAMPA"
            vntParam = Split(strParametri, ";")
            Call PrintBlaster((vntParam(0)), (vntParam(2)))
        Case Else
            bolRes = EstensioniRegoleBusiness(pobjAgt, strComando, strParametri)
    End Select
fine_mdiEstensioneAgenti:
    mdiEstensioneAgenti = bolRes
    
Exit Function
err_mdiEstensioneAgenti:
    Dim lngerr, strErr$
    bolRes = False
    lngerr = Err.Number: strErr = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("mdiEstensioneAgenti", lngerr, strErr))
Resume fine_mdiEstensioneAgenti

err_EseguiAzione:
    bolRes = False
    Call MXNU.MsgBoxEX(1348, vbCritical, 1007, "")
    On Local Error GoTo 0
Resume fine_mdiEstensioneAgenti

End Function

'Private Sub mnuModulo_Click(Index As Integer)
'    Call frmModuli.MostraSubMenu(mnuModulo(Index).tag, mnuModulo(Index).Caption)
'End Sub

'Private Sub tlbDesigner_ButtonClick(ByVal Button As MSComctlLib.Button)
'  Dim objManagedControl As Object
'  Dim i As Integer
'  Dim oDoc As MSXML2.DOMDocument
'  Dim oActiveNode As MSXML2.IXMLDOMNode
'  Dim bolRes As Boolean
'  Dim objSys As FileSystemObject
'  Dim objParentFold As Folder
'  Dim objChildFold As Folder
'  Dim strPathSave As String
'
'    Select Case LCase(Button.key)
'        Case "unvisible"
'            Call Designer.AssegnaProprieta(Button.key, Button.Value)
'            Designer.bolModify = True
'        Case "disable"
'            Call Designer.AssegnaProprieta(Button.key, Button.Value)
'            Designer.bolModify = True
'        Case "save"
'            With tlbDesigner.Buttons
'                .Item("unvisible").Value = tbrUnpressed
'                .Item("disable").Value = tbrUnpressed
'                .Item("showdiff").Value = tbrUnpressed
'                For i = 1 To .Count
'                    .Item(i).Enabled = False
'                Next i
'                'Rimane sempre attivo il bottone "showdiff"
'                .Item("showdiff").Enabled = True
'            End With
'            'Prima di tutto termino la modalità di editing...
'            '.. e ritrasformo i controlli di appoggio nei controlli standard Metodo (che erano invisibili)
'            Call Designer.TerminateEditMode
'            ' Genero file xml (nel path temporaneo in quanto passo false)
'            Call Designer.ExportDesign(False)
'            ' Termino il designer
'            Call Designer.TerminateDesigner
'
'            If Designer.bolSalvato Then
'                ' Salvataggio automatico ....Copio il file dal temporaneo alla versione definitiva
'                ' Lettura del file xml per vedere se vale per tutti gli utenti o per i singoli
'                Set oDoc = New MSXML2.DOMDocument
'                bolRes = oDoc.Load(MXNU.PercorsoDesigner & "\DesignSettings.xml")
'                If (bolRes) Then
'                    Set objSys = New FileSystemObject
'                    Set oActiveNode = oDoc.selectSingleNode("settings/firm[@name='" & MXNU.DittaAttiva & "']/version[@name='" & Designer.VersioneDesigner & "']")
'                    If Not (oActiveNode Is Nothing) Then
'                        If LCase(oActiveNode.Attributes.getNamedItem("typeversion").nodeValue) = "allusers" Then
'                            ' Salvataggio sulla versione
'                            With Designer
'                                Call objSys.CopyFile(MXNU.PercorsoDesigner & "\TempDesign\" & .strFileXML, MXNU.PercorsoDesigner & "\" & MXNU.DittaAttiva & "\" & .VersioneDesigner & "\" & .strFileXML, True)
'                            End With
'                        ElseIf LCase(oActiveNode.Attributes.getNamedItem("typeversion").nodeValue) = "singleusers" Then
'                            ' Salvataggio per ogni utente...
'                            Set objParentFold = objSys.GetFolder(MXNU.PercorsoDesigner & "\" & MXNU.DittaAttiva & "\" & Designer.VersioneDesigner)
'                            For Each objChildFold In objParentFold.SubFolders
'                                strPathSave = MXNU.PercorsoDesigner & "\" & MXNU.DittaAttiva & "\" & Designer.VersioneDesigner & "\" & objChildFold.Name
'                                Call objSys.CopyFile(MXNU.PercorsoDesigner & "\TempDesign\" & Designer.strFileXML, strPathSave & "\" & Designer.strFileXML, True)
'                            Next objChildFold
'                            Set objChildFold = Nothing
'                            Set objParentFold = Nothing
'                        End If
'                    End If
'                    Set oActiveNode = Nothing
'                    Set objSys = Nothing
'                End If
'                Set oDoc = Nothing
'            Else
'                ' La form seguente mi permette di decidere su quale versione copiare il file xml temporaneo
'                Load frmDefDesigner
'                frmDefDesigner.Show
'            End If
'            cmbDesignType.ListIndex = -1
'            cmbDesignType.Enabled = True
'        Case "discard"
'            With tlbDesigner.Buttons
'                With tlbDesigner.Buttons
'                    .Item("unvisible").Value = tbrUnpressed
'                    .Item("disable").Value = tbrUnpressed
'                    .Item("showdiff").Value = tbrUnpressed
'                    For i = 1 To .Count
'                        .Item(i).Enabled = False
'                    Next i
'                    'Rimane sempre attivo il bottone "showdiff"
'                    .Item("showdiff").Enabled = True
'                End With
'            End With
'            With Designer
'                .bolModify = False
'                Barra.Buttons(idxBottoneDesigner).Enabled = True
'                Call .TerminateEditMode
'                Call .TerminateDesigner
'                Set objManagedControl = FormAttiva.ActiveControl
'                Call .ApplyDesign
'                cmbDesignType.ListIndex = -1
'                cmbDesignType.Enabled = True
'                ' Faccio refresh perché ogni tanto si perde il cursore ....con refresh si sistema !!
'                DoEvents
'                FormAttiva.Refresh
'                DoEvents
'            End With
'
'        Case "showdiff"
'            cmbDesignType.Enabled = Not (Button.Value = tbrPressed)
'            Call Designer.ShowDiff((Button.Value = tbrPressed))
'            SchedaDesigner(0).Visible = (Button.Value = tbrPressed)
'        Case "undoctrl"
'            Call Designer.UndoCtrl
'        Case "setting"
'            Load FrmDesignerProp
'    End Select
'
'End Sub

'Private Sub tlbDesigner_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'Dim i As Integer
'  Dim oDoc As MSXML2.DOMDocument
'  Dim oActiveNode As MSXML2.IXMLDOMNode
'  Dim bolRes As Boolean
'  Dim objSys As FileSystemObject
'  Dim objParentFold As Folder
'  Dim objChildFold As Folder
'  Dim strPathSave As String
'
'    With tlbDesigner.Buttons
'        .Item("unvisible").Value = tbrUnpressed
'        .Item("disable").Value = tbrUnpressed
'        .Item("showdiff").Value = tbrUnpressed
'        For i = 1 To .Count
'            .Item(i).Enabled = False
'        Next i
'        .Item("showdiff").Enabled = True
'    End With
'
'    Call Designer.TerminateEditMode
'    Call Designer.ExportDesign(False)
'    Call Designer.TerminateDesigner
'
'    Select Case LCase(ButtonMenu.key)
'        Case "savedesign":
'            If Designer.bolSalvato Then
'                ' Ho già salvato almeno una volta quindi vado in sovrascrittura (senza richiesta di) dei file xml
'                Set oDoc = New MSXML2.DOMDocument
'                bolRes = oDoc.Load(MXNU.PercorsoDesigner & "\DesignSettings.xml")
'                If (bolRes) Then
'                    Set objSys = New FileSystemObject
'                    Set oActiveNode = oDoc.selectSingleNode("settings/firm[@name='" & MXNU.DittaAttiva & "']/version[@name='" & Designer.VersioneDesigner & "']")
'                    If Not (oActiveNode Is Nothing) Then
'                        If LCase(oActiveNode.Attributes.getNamedItem("typeversion").nodeValue) = "allusers" Then
'                            ' Salvataggio sulla versione...
'                            With Designer
'                                Call objSys.CopyFile(MXNU.PercorsoDesigner & "\TempDesign\" & .strFileXML, MXNU.PercorsoDesigner & "\" & MXNU.DittaAttiva & "\" & .VersioneDesigner & "\" & .strFileXML, True)
'                            End With
'                        ElseIf LCase(oActiveNode.Attributes.getNamedItem("typeversion").nodeValue) = "singleusers" Then
'                            ' Salvataggio per ogni utente...
'                            Set objParentFold = objSys.GetFolder(MXNU.PercorsoDesigner & "\" & MXNU.DittaAttiva & "\" & Designer.VersioneDesigner)
'                            For Each objChildFold In objParentFold.SubFolders
'                                strPathSave = MXNU.PercorsoDesigner & "\" & MXNU.DittaAttiva & "\" & Designer.VersioneDesigner & "\" & objChildFold.Name
'                                Call objSys.CopyFile(MXNU.PercorsoDesigner & "\TempDesign\" & Designer.strFileXML, strPathSave & "\" & Designer.strFileXML, True)
'                            Next objChildFold
'                            Set objChildFold = Nothing
'                            Set objParentFold = Nothing
'                        End If
'                        ' ... e cancello il temporaneo alla fine !!
'                        If objSys.FileExists(MXNU.PercorsoDesigner & "\TempDesign\" & Designer.strFileXML) Then
'                            Call objSys.DeleteFile(MXNU.PercorsoDesigner & "\TempDesign\" & Designer.strFileXML, True)
'                        End If
'                    End If
'                    Set oActiveNode = Nothing
'                    Set objSys = Nothing
'                End If
'                Set oDoc = Nothing
'            Else
'                ' Anche se salvo sulla versione corrente mi viene richiesto di indicarla perché è la prima volta
'                Load frmDefDesigner
'                frmDefDesigner.Show
'            End If
'        Case "saveversion":
'            ' Salva con nome ...
'            Call Designer.TerminateEditMode
'            Call Designer.ExportDesign(False)
'            Call Designer.TerminateDesigner
'            Load frmDefDesigner
'            frmDefDesigner.Show
'
'    End Select
'    cmbDesignType.ListIndex = -1
'    cmbDesignType.Enabled = True
'
'End Sub

'Private Sub tlbFile_ButtonClick(ByVal Button As MSComctlLib.Button)
'     Select Case LCase(Button.key)
'        Case "new"
'            Call MXQM.DoAction("nuovomodulo", "")
'        Case "open"
'            Call MXQM.DoAction("aprimodulo", "")
'        Case "save"
'            Call MXQM.DoAction("salvamodulo", "")
'        Case "preview"
'            Call MXQM.DoAction("anteprimastampa", "")
'        Case "print"
'            Call MXQM.DoAction("stampa", "")
'        Case "sign"
'            Call MXQM.DoAction("firma", "")
'        Case "deleterow"
'            Call MXQM.DoAction("delriga", "")
'        Case "insertrow"
'            Call MXQM.DoAction("insriga", "")
'        Case "appendrow"
'            Call MXQM.DoAction("accriga", "")
'        Case "delete"
'            Call MXQM.DoAction("deldoc", "")
'        Case "list"
'            Call MXQM.DoAction("elenco", "")
'        Case "query"
'            Call MXQM.DoAction("interroga", "")
'        Case "sortup"
'            Call MXQM.DoAction("sortup", "")
'        Case "sortdown"
'            Call MXQM.DoAction("sortdown", "")
'        Case "filter_write"
'            Call MXQM.DoAction("filter_write", "")
'        Case "filter_zot"
'            Call MXQM.DoAction("filter_zot", "")
'        Case "filter_clear"
'            Call MXQM.DoAction("filter_clear", "")
'        Case "execute"
'            Call MXQM.DoAction("execute", "")
'        Case Else
'    End Select
'
'End Sub

'Private Sub tlbFile_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'
'    Select Case LCase(ButtonMenu.key)
'        Case "new"
'            Call MXQM.DoAction("nuovomodulo", "")
'        Case "open"
'            Call MXQM.DoAction("aprimodulo", "")
'    End Select
'
'End Sub

Private Sub StatusBar_PaneClick(ByVal Pane As XtremeCommandBars.StatusBarPane)
    Select Case Pane.ID
        Case STATUSBAR_ID_PANELLINGUA: Call MostraPopUpLingue
        Case STATUSBAR_ID_PANELDESIGNER
            If ((MXNU.ControlloModuliChiave(modMyErpRunTime) = 0) _
                Or (MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0) _
                Or (MXNU.ControlloModuliChiave(modMyErpDesign) = 0)) Then
                
                Call frmpopup.MostraPopUpDesigner(Me, "Designer")
            End If
    End Select
End Sub

