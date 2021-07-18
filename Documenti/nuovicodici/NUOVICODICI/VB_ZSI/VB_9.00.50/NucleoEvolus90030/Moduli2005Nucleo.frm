VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "CODEJOCK.SHORTCUTBAR.UNICODE.V12.1.1.OCX"
Begin VB.Form frmModuli2005 
   BorderStyle     =   0  'None
   Caption         =   "Outlook Style Panels"
   ClientHeight    =   6615
   ClientLeft      =   2670
   ClientTop       =   4260
   ClientWidth     =   3615
   LinkTopic       =   "Form2"
   ScaleHeight     =   6615
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin XtremeShortcutBar.ShortcutBar ShortcutBar1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _Version        =   786433
      _ExtentX        =   6376
      _ExtentY        =   11668
      _StockProps     =   64
      VisualTheme     =   3
   End
   Begin MSComctlLib.ImageList imlShortcutBarIconsBig 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":0892
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":0FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":16BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":1DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":22A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":29B5
            Key             =   "Cartella"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":30C9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlShortcutBarIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":37DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":3B77
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":3E0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":424E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":44B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":4727
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":4B47
            Key             =   "Cartella"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moduli2005Nucleo.frx":4F52
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmModuli2005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private McolUserForms As Collection
Private NPannelli As Integer


Private MColShortcutItemResourcesID As New Collection
Public Sub CambiaRisorse()
    Dim Item As ShortcutBarItem
    Dim i%
    For i = 0 To NAVBAR_OFFSET - 1
        ShortcutBar1.Item(i).Caption = MXNU.CaricaStringaRes(MColShortcutItemResourcesID(CStr(ShortcutBar1.Item(i).Id)))
    Next i
End Sub


Private Sub Form_Load()
    Dim Item As ShortcutBarItem
    
    Call CambiaCharSet(Me) ' Aggiornamento multilingua (10/07/2007)
    
    If Not (McolFormsInNavBar Is Nothing) Then
        Set McolFormsInNavBar = Nothing
    End If
    Set McolFormsInNavBar = New Collection
        
    Load frmMenu    'modificata assegnazione della variabile frmModuli: altrimenti accedendo da metodo a frmMenu carica un'altra istanza
    Set frmModuli = frmMenu
    
    McolFormsInNavBar.Add Me.Name, Me.Name
    McolFormsInNavBar.Add frmModuli.Name, frmModuli.Name
    
    'Risorse in lingua per i componenti Codejock
    ShortcutBarGlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"
    
    ShortcutBar1.RemoveAllItems
    Set Item = ShortcutBar1.AddItem(ID_BAR_PROGMODULES, MXNU.CaricaStringaRes(11989), frmModuli.hwnd)     '"Moduli Programma"
    MColShortcutItemResourcesID.Add 11989, CStr(Item.Id)
   
    ShortcutBar1.ExpandedLinesCount = 5
    ShortcutBar1.Selected = ShortcutBar1.FindItem(ID_BAR_PROGMODULES)
    
    Dim strFile As String
    Dim strRiga As String
    Dim i As Integer
    Dim frmContainer As Form  'frmUserPanelContainer
    Dim NomeUserC As String, strTitolo As String
    Dim picSmall As StdPicture, picBig As StdPicture
    Dim idxIcon As Long
    
    Set McolUserForms = New Collection
    
    strFile = CercaDirFile("UserPanels.Ini", MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & ";" & MXNU.PercorsoPers & ";" & MXNU.PercorsoPgm)
    i = 1
    Do
        strRiga = MXNU.LeggiProfilo(strFile, "UserPanels", i, "")
        If strRiga <> "" Then
            NPannelli = NPannelli + 1
            NomeUserC = Split(strRiga, ",")(0)
            strTitolo = Split(strRiga, ",")(1)
            On Local Error Resume Next
            frmContainer.NomeWrapper = Split(strRiga, ",")(2)
            On Local Error GoTo 0
            Load frmContainer
            Set Item = ShortcutBar1.AddItem(NAVBAR_OFFSET + NPannelli, strTitolo, frmContainer.hwnd)
            'Affinchè l'icona venga visualizzata correttamente, l'indice all'interno degli imagelist deve corrispondere all'id utilizzato per aggiungere l'item alla shortcutbar
            If frmContainer.GetIcon(picSmall, picBig) Then
                imlShortcutBarIconsBig.ListImages.Add NAVBAR_OFFSET + NPannelli, , picBig
                imlShortcutBarIcons.ListImages.Add NAVBAR_OFFSET + NPannelli, , picSmall
            Else
                'Nel caso nello usercontrol non siano presenti le image con l'icona, per evitare che venga visualizzata un'icona casuale
                'aggiungo nuovamente agli imagelist l'icona standard della cartellina con l'indice corrispondente
                imlShortcutBarIconsBig.ListImages.Add NAVBAR_OFFSET + NPannelli, , imlShortcutBarIconsBig.ListImages.Item("Cartella").Picture
                imlShortcutBarIcons.ListImages.Add NAVBAR_OFFSET + NPannelli, , imlShortcutBarIcons.ListImages.Item("Cartella").Picture
            End If
            frmContainer.Hide
            'If NPannelli = 1 Then McolFormsInNavBar.Add frmContainer.Name, frmContainer.Name
            McolFormsInNavBar.Add frmContainer.Name, frmContainer.Name
            McolUserForms.Add frmContainer.hwnd
                        
            
            'ShortcutBar1.Selected = Item
            'DoEvents
            'Item.Visible = False
            ShortcutBar1.ExpandedLinesCount = ShortcutBar1.ExpandedLinesCount + 1
            Set frmContainer = Nothing
        End If
        i = i + 1
    Loop While strRiga <> ""
    
    ShortcutBar1.AddImageList imlShortcutBarIcons
    ShortcutBar1.AddImageList imlShortcutBarIconsBig
    
    ShortcutBar1.Selected = ShortcutBar1.FindItem(ID_BAR_PROGMODULES)
    DoEvents
    metodo.DockingPaneManager.RecalcLayout
End Sub


Private Sub Form_Resize()
    Me.ShortcutBar1.Height = Me.ScaleHeight
    Me.ShortcutBar1.Width = Me.ScaleWidth
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload frmMenu
    If Not (McolUserForms Is Nothing) Then
        Dim frmHWnd As Long, i As Long
        Dim bolTrovato As Boolean
        While McolUserForms.Count > 0
            frmHWnd = McolUserForms(1)
            bolTrovato = False
            For i = 0 To Forms.Count - 1
                If Forms(i).hwnd = frmHWnd Then
                    Unload Forms(i)
                    McolUserForms.Remove 1
                    bolTrovato = True
                    Exit For
                End If
            Next i
            If Not bolTrovato Then   'Se non trovo l'hwnd nella collection forms, allora la form è già stata scaricata
                McolUserForms.Remove 1
            End If
        Wend
        Set McolUserForms = Nothing
    End If
    Set MColShortcutItemResourcesID = Nothing
    Set frmModuli2005 = Nothing
End Sub

Private Sub ShortcutBar1_SelectedChanged(ByVal Item As XtremeShortcutBar.IShortcutBarItem)
    If Not metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR) Is Nothing Then
        metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Title = Item.Caption
        metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Select
    End If
    
    'per sistema di messaggistica
    If (Item.Id = ID_BAR_MESSAGES) Then
        Dim oPane As XtremeDockingPane.Pane
        Dim oPane1 As XtremeDockingPane.Pane
        Set oPane = metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR2)
        If (Not oPane Is Nothing) Then
            If (oPane.Closed Or oPane.Hidden) Then
                oPane.Close
                oPane.Select
                Set oPane1 = metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1)
                If Not (oPane1 Is Nothing) Then
                    If Not oPane1.Closed Then   'Và in errore se si tenta di fare l'attach su un pannello chiuso (vedi anche anomalia 9065)
                        oPane1.AttachTo oPane
                        oPane.Select
                        metodo.DockingPaneManager.AttachPane oPane, oPane1
                    End If
                End If
            End If
        End If
        Set oPane = Nothing
    End If
    metodo.DockingPaneManager.RecalcLayout
End Sub
