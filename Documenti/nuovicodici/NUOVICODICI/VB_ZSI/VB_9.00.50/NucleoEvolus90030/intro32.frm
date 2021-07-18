VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0CFFC63E-D04B-4629-B582-D4CB9B33E9D6}#1.0#0"; "MXKIT.OCX"
Object = "{C69DDE6E-C81F-4783-B0EB-7A3935551178}#1.0#0"; "MXCTRL.OCX"
Object = "{5EA404D6-3403-4E1A-AF53-5172578311FD}#1.0#0"; "MXBUSINESS.OCX"
Begin VB.Form frmIntro 
   BackColor       =   &H00632325&
   BorderStyle     =   0  'None
   ClientHeight    =   3870
   ClientLeft      =   4275
   ClientTop       =   3405
   ClientWidth     =   6705
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   6  'Inside Solid
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H80000008&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3870
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   975
      Left            =   1080
      TabIndex        =   8
      Top             =   1620
      Visible         =   0   'False
      Width           =   3135
      _Version        =   524288
      _ExtentX        =   5530
      _ExtentY        =   1720
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
      SpreadDesigner  =   "intro32.frx":0000
      AppearanceStyle =   0
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CrystalActiveXReportViewer1 
      Height          =   1635
      Left            =   2460
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   3615
      _cx             =   6376
      _cy             =   2884
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
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
   Begin MXCtrl.MWSchedaBox SchEvolus 
      Height          =   1155
      Left            =   3000
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2037
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LightColor      =   6908265
      ScaleWidth      =   1575
      ScaleHeight     =   1155
   End
   Begin MSComctlLib.ImageList imglist 
      Left            =   0
      Top             =   825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   384
      ImageHeight     =   288
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "intro32.frx":048F
            Key             =   "M98"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "intro32.frx":1D3BB
            Key             =   "MXP"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "intro32.frx":6DB0D
            Key             =   "MXP1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "intro32.frx":BEB5F
            Key             =   "M2005"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "intro32.frx":C6B90
            Key             =   "Evolus"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "intro32.frx":CA5B8
            Key             =   "Candidate"
         EndProperty
      EndProperty
   End
   Begin MXKit.CTLXKit CTLXKit1 
      Left            =   0
      Top             =   0
      _ExtentX        =   2249
      _ExtentY        =   1296
      _ExtentID       =   "847B0E1E"
   End
   Begin MXBusiness.CTLXBus CTLXBus1 
      Left            =   1290
      Top             =   0
      _ExtentX        =   2143
      _ExtentY        =   1296
      _ExtentID       =   "847B0E1E"
   End
   Begin VB.Label Label1 
      Caption         =   $"intro32.frx":CE9C0
      Height          =   615
      Left            =   1920
      TabIndex        =   7
      Top             =   3180
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Image ImgWin7Logo 
      Height          =   1080
      Left            =   120
      Picture         =   "intro32.frx":CEA49
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Tag             =   "Copyright"
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Metodo '98"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   180
      TabIndex        =   3
      Tag             =   "Product"
      Top             =   3000
      Width           =   1170
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1020
      TabIndex        =   2
      Tag             =   "Platform"
      Top             =   3000
      Width           =   570
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Platform"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2340
      TabIndex        =   1
      Tag             =   "Version"
      Top             =   3000
      Width           =   585
   End
   Begin VB.Label lblOperation 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Tag             =   "LicenseTo"
      Top             =   3660
      Width           =   3915
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

DefInt A-Z
Public NascondiFinestra As Boolean

Dim mChildFormRegion As Long
Private xp As Long, yp As Long

Public Sub MostraMessaggioOperazione(vntStringID As Variant, Optional vntParam As Variant)

    If IsNumeric(vntStringID) Then
        frmIntro.lblOperation.Caption = MXNU.CaricaStringaRes(vntStringID, vntParam)
    Else
        frmIntro.lblOperation.Caption = CStr(vntStringID)
    End If
    frmIntro.lblOperation.Refresh
    DoEvents

End Sub

Private Sub Form_Initialize()
    InitCommonControls   'Necessario per il corretto rendering degli oggetti in 3D col Manifest
End Sub

Private Sub Form_Load()
    Dim strPictureKey As String
    Dim intPos As Integer
    xp = Screen.TwipsPerPixelX
    yp = Screen.TwipsPerPixelY
    '*** FOR METODO XP ***
    #If ISMETODO2005 = 1 Then
        strPictureKey = "Evolus"
    #Else
        strPictureKey = "MXP"
    #End If
    lblPlatform.Caption = "per Microsoft Windows"
    #If ISMETODOXP = 1 Then
        'Me.Width = 5795
        Me.width = 5735
        Me.Height = 4335
        With lblProductName
            .ForeColor = vbBlack
            .Top = lblProductName.Top + 600
            .Left = lblProductName.Left + 150
        End With
        
        With lblCopyright
            .ForeColor = vbBlack
            .Top = lblCopyright.Top + 600
            .Left = lblCopyright.Left + 150
        End With
        
        With lblVersion
            .ForeColor = vbBlack
            .Top = lblVersion.Top + 600
            .Left = lblVersion.Left + 150
        End With
        
        With lblPlatform
            .ForeColor = vbBlack
            .Top = lblPlatform.Top + 600
            .Left = lblPlatform.Left + 150
            .Caption = "per Microsoft Windows"
        End With
        
        With lblOperation
            .ForeColor = &H47FF       'FF4700
            .Top = lblOperation.Top + 450
            .Left = lblOperation.Left + 150
            '.ForeColor = &H3399FF
        End With
        'smusso gli angoli della form
        #If ISMETODO2005 <> 1 Then
            mChildFormRegion = CreateRoundRectRgn(0, 0, Me.width / xp, Me.Height / yp, 40, 40)
            SetWindowRgn Me.hwnd, mChildFormRegion, False
        #Else
            Me.BackColor = vbBlack
        #End If
      
    #End If
    Me.Picture = imglist.ListImages(strPictureKey).Picture
    
    lblProductName.Caption = App.Title
    lblVersion.Caption = "Versione " & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
    
    lblCopyright.Caption = "Copyright © Metodo Spa. Tutti i diritti riservati."
    
    #If ISMETODO2005 = 1 Then
        lblProductName.ForeColor = vbBlack
        lblPlatform.ForeColor = vbBlack
        lblVersion.ForeColor = vbBlack
        lblCopyright.ForeColor = vbBlack
        lblVersion.Left = lblProductName.Left + lblProductName.width + 30
        lblPlatform.Left = lblVersion.Left + lblVersion.width + 30
        'Inizializzo i controlli, altrimenti in caso di login integrato le linguette non si disegnano correttamente
        Dim oM2005 As MXCtrl.M2005Setup
        Set oM2005 = New MXCtrl.M2005Setup
        oM2005.ISMETODO2005 = True
        Call CambiaSchemaColori(True)
        Set oM2005 = Nothing
        ImgWin7Logo.Visible = True
        ImgWin7Logo.ZOrder 0
    #End If
    
    If (Not NascondiFinestra) Then
        Call CentraFinestra(Me.hwnd)
        Me.Show
    End If

End Sub

Private Sub Form_Paint()
    #If ISMETODO2005 = 1 Then
        Line (0, 0)-(Me.ScaleWidth, 0), vbBlack
        Line (0, 0)-(0, Me.ScaleHeight), vbBlack
    #End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
#If ISMETODOXP = 1 Then
    'scarico gli oggetti che mi hanno smussato la form
    SetWindowRgn Me.hwnd, 0, False
    DeleteObject mChildFormRegion
#End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmIntro = Nothing
End Sub






