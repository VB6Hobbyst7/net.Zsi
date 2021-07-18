VERSION 5.00
Object = "{B50B6FC0-1ED7-4276-ADAE-2D69E0FC3E21}#1.0#0"; "MXCTRL.OCX"
Begin VB.Form frmInfoStp 
   Caption         =   "Informazioni di Stampa"
   ClientHeight    =   3750
   ClientLeft      =   2790
   ClientTop       =   4170
   ClientWidth     =   8685
   Icon            =   "INFOSTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   8685
   Begin MXCtrl.MWSchedaBox MWSchedaBox1 
      Height          =   3735
      Left            =   -60
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6588
      ForeColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LightColor      =   -2147483633
      ShadowColor     =   -2147483633
      ScaleWidth      =   8775
      ScaleHeight     =   3735
      Begin VB.Frame Frame2 
         Caption         =   "Filtro di Stampa"
         Height          =   1845
         Left            =   90
         TabIndex        =   2
         Top             =   0
         WhatsThisHelpID =   37669
         Width           =   8625
         Appearance      =   0  'Flat
         BorderStyle     =   1  'None
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Begin MXCtrl.MWSchedaBox MWSchedaBox2 
            Height          =   1515
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2672
            ForeColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LightColor      =   -2147483633
            ShadowColor     =   -2147483633
            ScaleWidth      =   8415
            ScaleHeight     =   1515
            Begin VB.TextBox Etcro 
               BackColor       =   &H80000005&
            Height          =   270
               Index           =   2
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   6
               top = 975
               Width           =   5685
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font
            Weight          =   400
            EndProperty  
            End
            Begin VB.TextBox Etcro 
               BackColor       =   &H80000005&
            Height          =   270
               Index           =   1
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   5
               top = 555
               Width           =   5685
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font
            Weight          =   400
            EndProperty  
            End
            Begin VB.TextBox Etcro 
               BackColor       =   &H80000005&
            Height          =   270
               Index           =   0
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   4
               top = 135
               Width           =   5685
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font
            Weight          =   400
            EndProperty  
            End
            Begin MXCtrl.MWEtichetta etc
            Height          =   270
               Index           =   0
               Left            =   60
               TabIndex        =   9
               Top             =   960
               WhatsThisHelpID =   10000
               Width           =   2415
               VariousPropertyBits=   19
               Caption         =   "Descrizione"
            Size            =   "2884;556"
               SpecialEffect   =   0
            FontHeight      =   225
               FontCharSet     =   0
               FontPitchAndFamily=   2
            BorderColor     =   -2147483624
            BorderStyle = 0
            FontName        =   "Tahoma"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name = "Tahoma"
            Size = 9
            Charset = 0
            Weight = 400
            Underline = 0 'False
            Italic = 0 'False
            Strikethrough = 0 'False
            EndProperty
            VAlign = 1
            BackColor = -2147483633
            ShadowColor = -2147483624
            LightColor = -2147483624
            End
            Begin MXCtrl.MWEtichetta etc
            Height          =   270
               Index           =   2
               Left            =   60
               TabIndex        =   8
               Top             =   120
               WhatsThisHelpID =   11178
               Width           =   2415
               VariousPropertyBits=   19
               Caption         =   "Nome"
            Size            =   "2884;556"
               SpecialEffect   =   0
            FontHeight      =   225
               FontCharSet     =   0
               FontPitchAndFamily=   2
            BorderColor     =   -2147483624
            BorderStyle = 0
            FontName        =   "Tahoma"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name = "Tahoma"
            Size = 9
            Charset = 0
            Weight = 400
            Underline = 0 'False
            Italic = 0 'False
            Strikethrough = 0 'False
            EndProperty
            VAlign = 1
            BackColor = -2147483633
            ShadowColor = -2147483624
            LightColor = -2147483624
            End
            Begin MXCtrl.MWEtichetta etc
            Height          =   270
               Index           =   4
               Left            =   60
               TabIndex        =   7
               Top             =   540
               WhatsThisHelpID =   11179
               Width           =   2415
               VariousPropertyBits=   19
               Caption         =   "Percorso"
            Size            =   "2884;556"
               SpecialEffect   =   0
            FontHeight      =   225
               FontCharSet     =   0
               FontPitchAndFamily=   2
            BorderColor     =   -2147483624
            BorderStyle = 0
            FontName        =   "Tahoma"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name = "Tahoma"
            Size = 9
            Charset = 0
            Weight = 400
            Underline = 0 'False
            Italic = 0 'False
            Strikethrough = 0 'False
            EndProperty
            VAlign = 1
            BackColor = -2147483633
            ShadowColor = -2147483624
            LightColor = -2147483624
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Stampa"
         Height          =   1755
         Left            =   90
         TabIndex        =   1
         Top             =   1980
         WhatsThisHelpID =   37670
         Width           =   8625
         Appearance      =   0  'Flat
         BorderStyle     =   1  'None
         BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8
         EndProperty
         Begin MXCtrl.MWSchedaBox MWSchedaBox3 
            Height          =   1395
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2461
            ForeColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LightColor      =   -2147483633
            ShadowColor     =   -2147483633
            ScaleWidth      =   8415
            ScaleHeight     =   1395
            Begin VB.TextBox Etcro 
               BackColor       =   &H80000005&
            Height          =   270
               Index           =   5
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   13
               top = 975
               Width           =   5685
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font
            Weight          =   400
            EndProperty  
            End
            Begin VB.TextBox Etcro 
               BackColor       =   &H80000005&
            Height          =   270
               Index           =   4
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   12
               top = 555
               Width           =   5685
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font
            Weight          =   400
            EndProperty  
            End
            Begin VB.TextBox Etcro 
               BackColor       =   &H80000005&
            Height          =   270
               Index           =   3
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   11
               top = 135
               Width           =   5685
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font
            Weight          =   400
            EndProperty  
            End
            Begin MXCtrl.MWEtichetta etc
            Height          =   270
               Index           =   1
               Left            =   60
               TabIndex        =   16
               Top             =   960
               WhatsThisHelpID =   10000
               Width           =   2415
               VariousPropertyBits=   19
               Caption         =   "Descrizione"
            Size            =   "2884;556"
               SpecialEffect   =   0
            FontHeight      =   225
               FontCharSet     =   0
               FontPitchAndFamily=   2
            BorderColor     =   -2147483624
            BorderStyle = 0
            FontName        =   "Tahoma"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name = "Tahoma"
            Size = 9
            Charset = 0
            Weight = 400
            Underline = 0 'False
            Italic = 0 'False
            Strikethrough = 0 'False
            EndProperty
            VAlign = 1
            BackColor = -2147483633
            ShadowColor = -2147483624
            LightColor = -2147483624
            End
            Begin MXCtrl.MWEtichetta etc
            Height          =   270
               Index           =   3
               Left            =   60
               TabIndex        =   15
               Top             =   120
               WhatsThisHelpID =   11325
               Width           =   2415
               VariousPropertyBits=   19
               Caption         =   "Nome"
            Size            =   "2884;556"
               SpecialEffect   =   0
            FontHeight      =   225
               FontCharSet     =   0
               FontPitchAndFamily=   2
            BorderColor     =   -2147483624
            BorderStyle = 0
            FontName        =   "Tahoma"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name = "Tahoma"
            Size = 9
            Charset = 0
            Weight = 400
            Underline = 0 'False
            Italic = 0 'False
            Strikethrough = 0 'False
            EndProperty
            VAlign = 1
            BackColor = -2147483633
            ShadowColor = -2147483624
            LightColor = -2147483624
            End
            Begin MXCtrl.MWEtichetta etc
            Height          =   270
               Index           =   5
               Left            =   60
               TabIndex        =   14
               Top             =   540
               WhatsThisHelpID =   11179
               Width           =   2415
               VariousPropertyBits=   19
               Caption         =   "Percorso"
            Size            =   "2884;556"
               SpecialEffect   =   0
            FontHeight      =   225
               FontCharSet     =   0
               FontPitchAndFamily=   2
            BorderColor     =   -2147483624
            BorderStyle = 0
            FontName        =   "Tahoma"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name = "Tahoma"
            Size = 9
            Charset = 0
            Weight = 400
            Underline = 0 'False
            Italic = 0 'False
            Strikethrough = 0 'False
            EndProperty
            VAlign = 1
            BackColor = -2147483633
            ShadowColor = -2147483624
            LightColor = -2147483624
            End
         End
      End
   End
End
Attribute VB_Name = "frmInfoStp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'##Form_Declarations
Public filtro_Nome As String
Public filtro_Percorso As String
Public filtro_Des As String
Public Stp_Nome As String
Public Stp_Percorso As String
Public Stp_Des As String

Private Sub Form_Load()
        
    Call MXNU.LeggiRisorseControlli(Me)
    Me.Caption = MXNU.CaricaStringaRes(23441)
    
    EtcRo(0).text = filtro_Nome
    EtcRo(1).text = filtro_Percorso
    EtcRo(2).text = filtro_Des

    EtcRo(3).text = Stp_Nome
    EtcRo(4).text = Stp_Percorso
    EtcRo(5).text = Stp_Des

    Call CentraFinestra(Me.hwnd)

End Sub

Private Sub Form_Paint()
    Call SchedaOmbreggiaControlli(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmInfoStp = Nothing

End Sub

Public Sub MostraImpstazioniStampa(ByVal cmbStampa As Control, vetNomiStampe() As String, ByVal objFiltro As MXKit.CFiltro, ByVal strNomeFiltro As String)
Dim vntItemData As Variant
Dim lngListIndex As Long

    On Local Error Resume Next
    'leggo i dati dal combo delle stampe
    With cmbStampa
        lngListIndex = .ListIndex
        vntItemData = .ItemData(lngListIndex)
    End With
    'leggo nome stampa
    Stp_Nome = Mid$(vetNomiStampe(vntItemData), InStrRev(vetNomiStampe(vntItemData), "\") + 1)
    'leggo descrizione stampa
    Stp_Des = cmbStampa.List(lngListIndex)
    'leggo percorso stampa
    Stp_Percorso = Left$(vetNomiStampe(vntItemData), InStrRev(vetNomiStampe(vntItemData), "\") - 1)
    Stp_Percorso = Replace(Stp_Percorso, "%PATHPERSDITTA%", MXNU.PercorsoPers & "\" & MXNU.DittaAttiva, , , vbTextCompare)
    Stp_Percorso = Replace(Stp_Percorso, "%PATHPERSDITTA-" & MXNU.DittaAttiva & "%", MXNU.PercorsoPers & "\" & MXNU.DittaAttiva, , , vbTextCompare)
    Stp_Percorso = Replace(Stp_Percorso, "%PATHPGM%", MXNU.PercorsoPgm & "\STAMPE", , , vbTextCompare)
    Stp_Percorso = Replace(Stp_Percorso, "%PATHPERS%", MXNU.PercorsoPers, , , vbTextCompare)
    'leggo descrizione filtro
    filtro_Des = objFiltro.DescrizioneFiltro
    'leggo percorso file filtro
    filtro_Percorso = CercaDirFile(strNomeFiltro & ".ss2", MXNU.PercorsoPers$ & "\" & MXNU.DittaAttiva & "\FILTRI;" & MXNU.PercorsoPers & "\FILTRI;" & MXNU.PercorsoFiltri)
    filtro_Nome = Mid$(filtro_Percorso, InStrRev(filtro_Percorso, "\") + 1)
    'leggo nome filtro
    filtro_Percorso = Left$(filtro_Percorso, InStrRev(filtro_Percorso, "\") - 1)
    On Local Error GoTo 0
'##Form_Load
    Me.Show vbModal
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'##Form_QueryUnload
End Sub
'##mResize_AfterResize

