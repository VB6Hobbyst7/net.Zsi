VERSION 5.00
Object = "{7EC174FF-49A5-4878-9AA0-74ED8D0C63DA}#1.0#0"; "mxctrl.ocx"
Begin VB.Form frmInfoStp 
   Caption         =   "Informazioni di Stampa"
   ClientHeight    =   3750
   ClientLeft      =   2790
   ClientTop       =   4170
   ClientWidth     =   8685
   Icon            =   "InfoStp.frx":0000
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
         Appearance      =   0  'Flat
         Caption         =   "Filtro di Stampa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   90
         TabIndex        =   2
         Top             =   0
         WhatsThisHelpID =   37669
         Width           =   8625
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
            FillWithGradient=   0   'False
            Begin VB.TextBox Etcro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   2
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   975
               Width           =   5685
            End
            Begin VB.TextBox Etcro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   1
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   555
               Width           =   5685
            End
            Begin VB.TextBox Etcro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   0
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   4
               Top             =   135
               Width           =   5685
            End
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   0
               Left            =   60
               Top             =   960
               WhatsThisHelpID =   10000
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
               Caption         =   "Descrizione"
            End
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   2
               Left            =   60
               Top             =   120
               WhatsThisHelpID =   11178
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
               Caption         =   "Nome"
            End
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   4
               Left            =   60
               Top             =   540
               WhatsThisHelpID =   11179
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
               Caption         =   "Percorso"
            End
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Stampa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   90
         TabIndex        =   1
         Top             =   1980
         WhatsThisHelpID =   37670
         Width           =   8625
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
            FillWithGradient=   0   'False
            Begin VB.TextBox Etcro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   5
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   975
               Width           =   5685
            End
            Begin VB.TextBox Etcro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   4
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   555
               Width           =   5685
            End
            Begin VB.TextBox Etcro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   3
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   135
               Width           =   5685
            End
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   1
               Left            =   60
               Top             =   960
               WhatsThisHelpID =   10000
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
               Caption         =   "Descrizione"
            End
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   3
               Left            =   60
               Top             =   120
               WhatsThisHelpID =   11325
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
               Caption         =   "Nome"
            End
            Begin MXCtrl.MWEtichetta etc 
               Height          =   270
               Index           =   5
               Left            =   60
               Top             =   540
               WhatsThisHelpID =   11179
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
               Caption         =   "Percorso"
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

#If IsMetodo2005 Then
    'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1
#End If
Public filtro_Nome As String
Public filtro_Percorso As String
Public filtro_Des As String
Public Stp_Nome As String
Public Stp_Percorso As String
Public Stp_Des As String

Private Sub Form_Load()
        
    Call MXNU.LeggiRisorseControlli(Me)
    Me.Caption = MXNU.CaricaStringaRes(23441)
    
    etcro(0).Text = filtro_Nome
    etcro(1).Text = filtro_Percorso
    etcro(2).Text = filtro_Des

    etcro(3).Text = Stp_Nome
    etcro(4).Text = Stp_Percorso
    etcro(5).Text = Stp_Des

    Call CentraFinestra(Me.hWnd)

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
    Stp_Des = cmbStampa.LIST(lngListIndex)
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
#If IsMetodo2005 Then
    'Inzializzazione Form per Metodo Evolus
    Call CambiaColoriControlli(Me)   'Da Mettere PRIMA di mResize
    On Local Error Resume Next
    Set mResize = New MxResizer.ResizerEngine
    If (Not mResize Is Nothing) Then
            Call mResize.Initialize(Me, , , , , 0, True, metodo, MXNU)
    End If
#End If
Call CentraFinestra(Me.hWnd)
Call CambiaCharSet(Me)
On Local Error GoTo 0
    Me.Show vbModal
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
#If IsMetodo2005 Then
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

#If IsMetodo2005 Then
    'Per Metodo Evolus
    Private Sub mResize_AfterResize()
        Call AvvicinaLing(Me)
    End Sub
#End If
Private Sub MWSchedaBox2_Click()

End Sub


Private Sub MWSchedaBox2_Paint()
    SchedaOmbreggiaControlli MWSchedaBox2
End Sub


Private Sub MWSchedaBox3_Paint()
    SchedaOmbreggiaControlli MWSchedaBox3
End Sub


