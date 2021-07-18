VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Begin VB.Form FrmAnteprima 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Anteprima di stampa"
   ClientHeight    =   4485
   ClientLeft      =   2370
   ClientTop       =   2865
   ClientWidth     =   7860
   FillStyle       =   0  'Solid
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
   Icon            =   "antepr32.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4485
   ScaleWidth      =   7860
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CrViewer91 
      Height          =   4395
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7755
      _cx             =   13679
      _cy             =   7752
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1040
   End
End
Attribute VB_Name = "FrmAnteprima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Public InElaborazione As Boolean

Private Const VK_LSHIFT = &HA0
Private Const VK_RSHIFT = &HA1

Public crwRpt As Object
Public MstrOriginalPrinter As String

Public Function AzioniMetodo(setAzione As enmFunzioniMetodo, Optional varparametro As Variant) As Variant

End Function

Private Sub CRViewer91_PrintButtonClicked(UseDefault As Boolean)
    If GetAsyncKeyState(VK_LSHIFT) < 0 Or GetAsyncKeyState(VK_RSHIFT) < 0 Then
        Call MXCREP.ImpostaStampanteViewer(CrViewer91, Me.hwnd)
        UseDefault = False
    Else
        UseDefault = True
    End If
End Sub

Private Sub CRViewer91_StopButtonClicked(ByVal loadingType As CrystalActiveXReportViewerLib11Ctl.CRLoadingType, UseDefault As Boolean)
    UseDefault = (loadingType <> crLoadingPages)
End Sub

Private Sub CrystalActiveXReportViewer1_CloseButtonClicked(UseDefault As Boolean)

End Sub

Private Sub Form_Activate()
    Call MXNU.ImpostaFormAttiva(Me)
End Sub

Private Sub Form_Load()
    Me.Left = 60
    Me.Top = 30
    'RIF. AN. #10169
    If metodo.WindowState = 1 Then metodo.WindowState = 2
    Me.Width = metodo.ScaleWidth - 120
    Me.Height = metodo.ScaleHeight - 60
    Me.BackColor = vbWindowBackground
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Cancel = MXCREP.StampaInCorso()
    'Cancel = InElaborazione
    Cancel = CrViewer91.IsBusy
End Sub

Private Sub Form_Resize()
    CrViewer91.Top = 0
    CrViewer91.Left = 0
    CrViewer91.Height = ScaleHeight
    CrViewer91.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GTestMode Then
        Sleep 1000
    End If
    Set crwRpt = Nothing
    MXCREP.IstanzeStampa = MXCREP.IstanzeStampa - 1
    'Anomalia 11465 (spostato qui il ripristino della stampante di default, altrimenti se si stampa dall'anteprima non presenta la stampante corretta)
    If MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "METODOW", "UsaAPIStampanti", 0) = 0 Then
        'If MstrOriginalPrinter <> "" Then Call SettaDefaultPrinter(MstrOriginalPrinter)
        If MstrOriginalPrinter <> "" Then Call MXCREP.SettaStampanteSistema(MstrOriginalPrinter)
    End If
    Set FrmAnteprima = Nothing
End Sub

