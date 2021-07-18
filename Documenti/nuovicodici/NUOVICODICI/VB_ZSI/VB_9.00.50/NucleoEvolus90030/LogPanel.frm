VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7EC174FF-49A5-4878-9AA0-74ED8D0C63DA}#1.0#0"; "MXCTRL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "CODEJOCK.COMMANDBARS.V12.0.0.OCX"
Begin VB.Form frmLogPanel 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   5235
   ClientTop       =   5385
   ClientWidth     =   9540
   Icon            =   "LogPanel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   4620
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MXCtrl.MWLog objLog 
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6694
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListLog 
      Left            =   780
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LogPanel.frx":000C
            Key             =   "Save"
            Object.Tag             =   "1070"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LogPanel.frx":05A6
            Key             =   "Print"
            Object.Tag             =   "1071"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LogPanel.frx":0B40
            Key             =   "Mail"
            Object.Tag             =   "1072"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   0
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmLogPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ID_LOG_SAVE = 1070
Private Const ID_LOG_PRINT = 1071
Private Const ID_LOG_MAIL = 1072

Private Sub InviaMail()
    'Dim frmMail As New frmLogMail
    Dim strFile As String
    Dim objMain As Object
    Dim objOutLook As Object

    Set objMain = CreateObject("MxOutlook.cMain") 'New MxOutlook.cMain
    Set objOutLook = objMain.CreaOutlook(MXNU.PercorsoPreferenze, MXNU.UtenteSistema, MXNU.PercorsoPers, MXNU.DittaAttiva)
    If Not (objOutLook Is Nothing) Then
        'Set frmMail.objOutLook = objOutLook
        'frmMail.Show 1
        
        'If Not frmMail.bolAnnullato Then
            strFile = MXNU.GetTempDir() & "MWLog" & MXNU.NTerminale & ".txt"
            If objLog.ExportLogFile(strFile) Then
                If objOutLook.Logon Then
                    'Con il primo parametro vuoto, uscirà la form di composizione messaggio del client di posta in uso
                    Call objOutLook.SendMailMapi("", "", "", strFile, True, True)
                End If
            End If
        'End If
        'Unload frmMail
    End If

END_InviaMail:
    On Local Error GoTo 0
    If Not (objOutLook Is Nothing) Then
        objOutLook.Logoff
        Set objOutLook = Nothing
    End If
    If Not (objMain Is Nothing) Then
        Set objMain = Nothing
    End If
    'Set frmMail = Nothing

End Sub

'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
'           FUNZIONI PRIVATE DELLA FORM
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
Private Sub StampaLOG()
#If ISM98SERVER = 0 Then

Dim objCRW As MXKit.CCrw
Dim strTitoloTesta As String
Dim strTitoloStampa As String
Dim strTitolo As String

    Screen.MousePointer = vbHourglass
    Set objCRW = MXCREP.CreaCCrw()
    objCRW.ClearOpzioniStp
    objCRW.OpzioniForm = STP_FILTRO
    Call objCRW.Stampante.LeggiVBPrinter
    Screen.MousePointer = vbDefault
    Call objCRW.MostraFrmStampa
    strTitolo = Me.Caption
    strTitoloStampa = strTitolo & " - " & MXNU.UtenteAttivo
    If Not objCRW.Stampa_Annullata Then
        If (objCRW.Periferica = "Stampante") Then
            Call objLog.PrintLog(strTitolo, objCRW)
        End If
    End If

fine_Stampa:
    Screen.MousePointer = vbDefault
    Set objCRW = Nothing
    Exit Sub
#End If
End Sub

Private Function EsportaLOG(ByVal strFileName As String) As Boolean
    EsportaLOG = objLog.ExportLogFile(strFileName)
    If (EsportaLOG) Then MXNU.MsgBoxEX 2080, vbInformation, "", Array(strFileName)
End Function

Private Function GetLOGName(strFileName As String) As Boolean
    GetLOGName = True
    strFileName = MXNU.GetTempDir() & "MWLog" & MXNU.NTerminale & ".txt"
    On Local Error GoTo err_GetLOGName
    With Cdlg
        .DialogTitle = MXNU.CaricaStringaRes(24343)
        .Filename = strFileName
        .Filter = "*.txt"
        .DefaultExt = "txt"
        .CancelError = True
        .Action = 2
        If (.Filename <> "") Then strFileName = .Filename
    End With

fine_GetLogName:
    On Local Error GoTo 0
    Exit Function

err_GetLOGName:
    GetLOGName = False
    Resume fine_GetLogName
End Function



Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case ID_LOG_SAVE
            Dim strFile As String
            If (GetLOGName(strFile)) Then
                Call EsportaLOG(strFile)
            End If
        Case ID_LOG_PRINT: Call StampaLOG
        Case ID_LOG_MAIL: Call InviaMail
    End Select
End Sub



Private Sub Form_Load()
    Dim tlb As XtremeCommandBars.CommandBar
    Dim Btn As XtremeCommandBars.CommandBarButton
    
    CommandBars.GlobalSettings.App = App
    CommandBars.EnableCustomization False
    CommandBars.AddImageList ImgListLog
    CommandBars.VisualTheme = metodo.CommandBars.VisualTheme
    
    Set tlb = CommandBars.Add("BarraLog", xtpBarTop)
    With tlb.Controls
        Set Btn = .Add(xtpControlButton, ID_LOG_SAVE, MXNU.CaricaStringaRes(25050))
        Btn.Style = xtpButtonIconAndCaption
        Set Btn = .Add(xtpControlButton, ID_LOG_PRINT, MXNU.CaricaStringaRes(25011))
        Btn.Style = xtpButtonIconAndCaption
        Set Btn = .Add(xtpControlButton, ID_LOG_MAIL, MXNU.CaricaStringaRes(10011))
        Btn.Style = xtpButtonIconAndCaption
    End With
    tlb.ContextMenuPresent = False
    tlb.ModifyStyle XTP_CBRS_GRIPPER, 0
    CommandBars.ActiveMenuBar.Visible = False
    CommandBars.Options.ShowExpandButtonAlways = False
    
    Call CambiaColoriControlli(Me)
    objLog.BackColor = Me.BackColor
    CommandBars.RecalcLayout
    
End Sub


Private Sub Form_Paint()
    Call SchedaOmbreggiaControlli(Me)
End Sub

Private Sub Form_Resize()
    objLog.Height = Me.Height - objLog.Top
    objLog.width = Me.width
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmLogPanel = Nothing
End Sub

