VERSION 5.00
Object = "{3CAD74A8-F6F9-47C0-BC9A-07A5F5BD8826}#1.0#0"; "MXCtrl.ocx"
Begin VB.Form FrmVisUtMod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   3345
   ClientTop       =   3570
   ClientWidth     =   7110
   Icon            =   "FrmVisUtMod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin MXCtrl.MWSchedaBox MWSchedaBox1 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7105
      _ExtentX        =   12541
      _ExtentY        =   6324
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleWidth      =   7110
      ScaleHeight     =   3585
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ClipControls    =   0   'False
         Height          =   780
         Left            =   600
         Picture         =   "FrmVisUtMod.frx":0442
         ScaleHeight     =   720
         ScaleMode       =   0  'User
         ScaleWidth      =   720
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   780
      End
      Begin VB.Image ImgXP 
         Height          =   780
         Left            =   600
         Picture         =   "FrmVisUtMod.frx":210C
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   780
      End
      Begin MXCtrl.MWEtichetta etcro 
         Height          =   300
         Index           =   4
         Left            =   4080
         Top             =   2760
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LightColor      =   12164479
         ShadowColor     =   12164479
         VAlign          =   1
      End
      Begin MXCtrl.MWEtichetta etcro 
         Height          =   300
         Index           =   3
         Left            =   2880
         Top             =   2760
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LightColor      =   12164479
         ShadowColor     =   12164479
         VAlign          =   1
      End
      Begin MXCtrl.MWEtichetta etcro 
         Height          =   300
         Index           =   2
         Left            =   2880
         Top             =   2280
         Width           =   2115
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LightColor      =   12164479
         ShadowColor     =   12164479
         VAlign          =   1
      End
      Begin MXCtrl.MWEtichetta etcro 
         Height          =   300
         Index           =   1
         Left            =   2880
         Top             =   1800
         Width           =   3855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LightColor      =   12164479
         ShadowColor     =   12164479
         VAlign          =   1
      End
      Begin MXCtrl.MWEtichetta etcro 
         Height          =   300
         Index           =   0
         Left            =   2880
         Top             =   1320
         Width           =   3855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LightColor      =   12164479
         ShadowColor     =   12164479
         VAlign          =   1
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   3
         Left            =   360
         Top             =   2760
         WhatsThisHelpID =   10195
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
         Caption         =   "Utente Ultima Modifica"
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   2
         Left            =   360
         Top             =   1320
         WhatsThisHelpID =   10192
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
         Caption         =   "Nome Tabella"
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   1
         Left            =   360
         Top             =   2280
         WhatsThisHelpID =   10194
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
         Caption         =   "Data Ultima Modifica"
      End
      Begin MXCtrl.MWEtichetta etc 
         Height          =   270
         Index           =   0
         Left            =   360
         Top             =   1800
         WhatsThisHelpID =   10193
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
         Caption         =   "Record Corrente"
      End
      Begin VB.Label lblVersion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Version"
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
         Left            =   1860
         TabIndex        =   3
         Tag             =   "Version"
         Top             =   720
         Width           =   3885
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   1860
         TabIndex        =   2
         Tag             =   "Application Title"
         Top             =   240
         Width           =   3885
      End
   End
End
Attribute VB_Name = "FrmVisUtMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mStrNomeTabella As String
Public MstrWHE As String
Public MstrDesRecord As String

Private Sub Form_Load()
    Dim intq As Integer, strSQL As String, hDY As CRecordSet

    Call MXNU.LeggiRisorseControlli(Me)
    Me.Caption = MXNU.CaricaStringaRes(23020)
    
    #If ISMETODOXP = 1 Then
        picIcon.Visible = False
        ImgXP.Visible = True
    #End If
    lblTitle.Caption = App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    EtcRo(0).Caption = mStrNomeTabella
    EtcRo(1).Caption = MstrDesRecord
    
    'rif.sch. A4519 - inserito l'alias della tabella per il campo di where
    ' strSQL = "SELECT t.UtenteModifica,t.DataModifica,u.Descrizione FROM " & mStrNomeTabella & " t LEFT OUTER JOIN TabUtenti u ON u.UserID=t.UtenteModifica WHERE " & MstrWHE
    strSQL = "SELECT t.UtenteModifica,t.DataModifica,u.Descrizione FROM " & mStrNomeTabella & " t LEFT OUTER JOIN TabUtenti u ON u.UserID=t.UtenteModifica WHERE t." & MstrWHE
    Set hDY = MXDB.dbCreaDY(hndDBArchivi, strSQL, TIPO_TABELLA)
    If Not MXDB.dbFineTab(hDY, TIPO_DYNASET) Then
        EtcRo(2).Caption = MXDB.dbGetCampo(hDY, TIPO_DYNASET, "DataModifica", "")
        EtcRo(3).Caption = MXDB.dbGetCampo(hDY, TIPO_DYNASET, "UtenteModifica", "")
        EtcRo(4).Caption = MXDB.dbGetCampo(hDY, TIPO_DYNASET, "Descrizione", "")
    End If
    intq = MXDB.dbChiudiDY(hDY)
    
    Call CentraFinestra(Me.hwnd)

'##Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set FrmVisUtMod = Nothing

End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'##Form_QueryUnload
End Sub
'##mResize_AfterResize

