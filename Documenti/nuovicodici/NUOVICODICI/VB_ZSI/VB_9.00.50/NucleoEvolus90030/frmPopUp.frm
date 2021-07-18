VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "CODEJOCK.COMMANDBARS.V13.0.0.OCX"
Begin VB.Form frmpopup 
   Caption         =   "PopUp"
   ClientHeight    =   2655
   ClientLeft      =   7020
   ClientTop       =   3360
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4350
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   0
      Top             =   0
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin VB.Menu GestDocR 
      Caption         =   "GestDocR"
      Visible         =   0   'False
      Begin VB.Menu GestDocRItem 
         Caption         =   "Inserisci"
         HelpContextID   =   40001
         Index           =   0
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "Cancella"
         HelpContextID   =   40002
         Index           =   1
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "Duplica"
         HelpContextID   =   40003
         Index           =   3
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "Taglia"
         HelpContextID   =   40004
         Index           =   5
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "Copia"
         HelpContextID   =   40005
         Index           =   6
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "Incolla"
         HelpContextID   =   40006
         Index           =   7
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "Adatta Colonna"
         HelpContextID   =   40007
         Index           =   9
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "Adatta Foglio"
         HelpContextID   =   40008
         Index           =   10
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "Nascondi Colonna"
         HelpContextID   =   40048
         Index           =   12
      End
      Begin VB.Menu GestDocRItem 
         Caption         =   "Mostra Colonna"
         HelpContextID   =   40049
         Index           =   13
         Begin VB.Menu GestDocRItemHide 
            Caption         =   "-"
            Index           =   0
         End
      End
   End
   Begin VB.Menu GestDocT 
      Caption         =   "GestDocT"
      Visible         =   0   'False
      Begin VB.Menu GestDocTItem 
         Caption         =   "Adatta Colonna"
         HelpContextID   =   40007
         Index           =   0
      End
      Begin VB.Menu GestDocTItem 
         Caption         =   "Adatta Foglio"
         HelpContextID   =   40008
         Index           =   1
      End
   End
   Begin VB.Menu mnuECR 
      Caption         =   "mnuECR"
      Visible         =   0   'False
      Begin VB.Menu mnuECRItem 
         Caption         =   "Item1"
         Index           =   0
      End
   End
   Begin VB.Menu MnuVert 
      Caption         =   "Verticale"
      Visible         =   0   'False
      Begin VB.Menu MnuVertItem 
         Caption         =   "Aggiungi &Gruppo"
         Index           =   0
      End
      Begin VB.Menu MnuVertItem 
         Caption         =   "R&inomina Gruppo Corrente"
         Index           =   1
      End
      Begin VB.Menu MnuVertItem 
         Caption         =   "&Rimuovi Ultimo Gruppo"
         Index           =   2
      End
      Begin VB.Menu MnuVertItem 
         Caption         =   "Rimuovi Ultimo &Elemento"
         Index           =   3
      End
   End
   Begin VB.Menu MnuVertGruppi 
      Caption         =   "Verticale"
      Visible         =   0   'False
      Begin VB.Menu MnuVertGruppiItem 
         Caption         =   "Aggiungi &Gruppo"
         Index           =   0
      End
      Begin VB.Menu MnuVertGruppiItem 
         Caption         =   "&Rimuovi Ultimo Gruppo"
         Index           =   1
      End
      Begin VB.Menu MnuVertGruppiItem 
         Caption         =   "R&inomina Gruppo Corrente"
         Index           =   2
      End
   End
   Begin VB.Menu MnuVertFigli 
      Caption         =   "Verticale"
      Visible         =   0   'False
      Begin VB.Menu MnuVertFigliItem 
         Caption         =   "Rimuovi & Elemento"
         Index           =   0
      End
      Begin VB.Menu MnuVertFigliItem 
         Caption         =   "Rimuovi  &Ultimo Elemento"
         Index           =   1
      End
      Begin VB.Menu MnuVertFigliItem 
         Caption         =   "&Rinomina Elemento"
         Index           =   2
      End
   End
   Begin VB.Menu mnuNodiPrevCommCli 
      Caption         =   "PreventivoCommesse"
      Visible         =   0   'False
      Begin VB.Menu mnuNodiPrevCommCliItem 
         Caption         =   "Crea nodo di pari livello..."
         Index           =   0
      End
      Begin VB.Menu mnuNodiPrevCommCliItem 
         Caption         =   "Crea nodo figlio..."
         Index           =   1
      End
      Begin VB.Menu mnuNodiPrevCommCliItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuNodiPrevCommCliItem 
         Caption         =   "Modifica..."
         Index           =   3
      End
      Begin VB.Menu mnuNodiPrevCommCliItem 
         Caption         =   "Modifica formula..."
         Index           =   4
      End
      Begin VB.Menu mnuNodiPrevCommCliItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuNodiPrevCommCliItem 
         Caption         =   "Imposta nodo Excel..."
         Index           =   6
      End
      Begin VB.Menu mnuNodiPrevCommCliItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuNodiPrevCommCliItem 
         Caption         =   "Elimina nodo"
         Index           =   8
      End
   End
   Begin VB.Menu Lingue 
      Caption         =   "Lingue"
      Visible         =   0   'False
      Begin VB.Menu LingueItem 
         Caption         =   "IT"
         Index           =   0
      End
   End
   Begin VB.Menu AllInOne 
      Caption         =   "AllInOne"
      Visible         =   0   'False
      Begin VB.Menu AllInOneItem 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Designer 
      Caption         =   "Designer"
      Visible         =   0   'False
      Begin VB.Menu DesignerItem 
         Caption         =   "Standard"
         Index           =   0
      End
   End
   Begin VB.Menu mnuSincStruttCC 
      Caption         =   "Sincronizza"
      Visible         =   0   'False
      Begin VB.Menu mnuSincStruttCCItem 
         Caption         =   "Ripristina"
         Index           =   0
      End
   End
   Begin VB.Menu mnuSSCC 
      Caption         =   "SSCC"
      Visible         =   0   'False
      Begin VB.Menu mnuSSCCItem 
         Caption         =   "Nuova Unità Logistica"
         Index           =   0
      End
      Begin VB.Menu mnuSSCCItem 
         Caption         =   "Seleziona Unità Logistica"
         Index           =   1
      End
   End
   Begin VB.Menu mnuSSCCPallet 
      Caption         =   "SSCCPallet"
      Visible         =   0   'False
      Begin VB.Menu mnuSSCCPalletItem 
         Caption         =   "Elimina Unità Logistica"
         Index           =   0
      End
      Begin VB.Menu mnuSSCCPalletItem 
         Caption         =   "Svuota Unità Logistica"
         Index           =   1
      End
   End
   Begin VB.Menu mnuSSCCProduct 
      Caption         =   "SSCCProduct"
      Visible         =   0   'False
      Begin VB.Menu mnuSSCCProductItem 
         Caption         =   "Elimina Prodotto"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmpopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private MidxMenu As Integer
'Menu Verticale

Public mnuVerticaleSel As String
Public indVerticaleSel As Integer
Public mColColonneNascoste As Collection
Public mStrMostraCol As String
Public mBolPersRighe As Boolean

Public Function MostraPopUp(pForm As Form, ByVal pstrNome As String) As Integer

    Dim mnu As Menu
    Dim idx As Integer
    Dim PopupBar As CommandBar
    Dim Btn As CommandBarControl
    Dim BtnMostra As CommandBarControl
    
    MidxMenu = -1
    
    On Local Error GoTo err_MostraPopUp
    'Anomalia nr. 10385
    If (pstrNome = "GestDocR") Then
        Set PopupBar = CommandBars.Add("GestDocR", xtpBarPopup)
        With PopupBar.Controls
            Set Btn = .Add(xtpControlButton, 50001, MXNU.CaricaStringaRes(40001))
            Set Btn = .Add(xtpControlButton, 50002, MXNU.CaricaStringaRes(40002))

            Set Btn = .Add(xtpControlButton, 50004, MXNU.CaricaStringaRes(40003))
            Btn.BeginGroup = True
            Set Btn = .Add(xtpControlButton, 50006, MXNU.CaricaStringaRes(40004))
            Btn.BeginGroup = True
            Set Btn = .Add(xtpControlButton, 50007, MXNU.CaricaStringaRes(40005))
            Set Btn = .Add(xtpControlButton, 50008, MXNU.CaricaStringaRes(40006))

            Set Btn = .Add(xtpControlButton, 50010, MXNU.CaricaStringaRes(40007))
            Btn.BeginGroup = True
            Set Btn = .Add(xtpControlButton, 50011, MXNU.CaricaStringaRes(40008))
            If mBolPersRighe Then

                Set Btn = .Add(xtpControlButton, 50013, MXNU.CaricaStringaRes(40048))
                Btn.BeginGroup = True
                Set BtnMostra = .Add(xtpControlPopup, 50014, MXNU.CaricaStringaRes(40049))
                If Not (mColColonneNascoste Is Nothing) Then
                    If (mColColonneNascoste.Count > 0) Then
                        For idx = 1 To mColColonneNascoste.Count
                            Call BtnMostra.CommandBar.Controls.Add(xtpControlButton, 50014 + idx, Split(mColColonneNascoste(idx), "|")(0))
                        Next idx
                    End If
                End If
            End If
            mStrMostraCol = ""
            PopupBar.ShowPopup
        End With
    Else
        Me.Controls(pstrNome).Visible = True
        For Each mnu In Me.Controls(pstrNome & "Item")
            If mnu.HelpContextID <> 0 Then
                mnu.Caption = MXNU.CaricaStringaRes(mnu.HelpContextID)
            End If
        Next
        mStrMostraCol = ""
        pForm.PopupMenu Me.Controls(pstrNome), vbPopupMenuLeftAlign
    End If
    
    On Local Error GoTo 0
fine_mostraPopUp:
    MostraPopUp = MidxMenu
    
    Unload Me
Exit Function

err_MostraPopUp:
    
    Call MXNU.MsgBoxEX(Err.Number & " - " & Err.Description, vbCritical, 1007)
    On Local Error GoTo 0
    
    Resume fine_mostraPopUp
End Function

Public Function MostraPopUpStrutturaCommesse(pForm As Form, ByVal pstrNome As String) As Integer

    Dim mnu As Menu
    Dim idx As Integer
    
    MidxMenu = -1
    
    On Local Error GoTo err_MostraPopUp
    Me.Controls(pstrNome).Visible = True
    For Each mnu In Me.Controls(pstrNome & "Item")
        If mnu.HelpContextID <> 0 Then
            mnu.Caption = MXNU.CaricaStringaRes(mnu.HelpContextID)
        End If
    Next
    
    Select Case UCase(pForm.Name)
        Case "FRMPREVCOMMCLI"
        Case "FRMSTRUTTSTANDARDCOMMCLI", "FRMANAGRAFICACOMMESSE"
            mnuNodiPrevCommCliItem(4).Enabled = False
            mnuNodiPrevCommCliItem(6).Enabled = False
    End Select

    pForm.PopupMenu Me.Controls(pstrNome), vbPopupMenuLeftAlign
        
    On Local Error GoTo 0
fine_mostraPopUp:
    MostraPopUpStrutturaCommesse = MidxMenu
    
    Unload Me
Exit Function

err_MostraPopUp:
    
    Call MXNU.MsgBoxEX(Err.Number & " - " & Err.Description, vbCritical, 1007)
    On Local Error GoTo 0
    
    Resume fine_mostraPopUp
End Function

Public Function MostraPopUpLingue(pForm As Form, ByVal pstrNome As String) As Integer
    Dim mnu As Menu
    Dim idx As Integer
    Dim strFile As String
    Dim strLingua As String
    Dim i As Integer
    
    MidxMenu = -1
    
    strFile = Dir(MXNU.PercorsoFileLocali & "\MWRES??.DAT")
    i = 1
    While strFile <> ""
        strLingua = Mid(strFile, 6, 2)
        If strLingua <> "IT" Then
            Load LingueItem(i)
            LingueItem(i).Caption = strLingua
            If MXNU.LinguaAttiva = strLingua Then
                Me.Controls(pstrNome & "Item")(i).Checked = True
            Else
                Me.Controls(pstrNome & "Item")(i).Checked = False
            End If
        Else
            If MXNU.LinguaAttiva = strLingua Then
                Me.Controls(pstrNome & "Item")(0).Checked = True
             Else
                Me.Controls(pstrNome & "Item")(0).Checked = False
            End If
        End If
        
        strFile = Dir()
        i = i + 1
    Wend
    
    On Local Error GoTo err_MostraPopUp
    Me.Controls(pstrNome).Visible = True
    For Each mnu In Me.Controls(pstrNome & "Item")
        If mnu.HelpContextID <> 0 Then
            mnu.Caption = MXNU.CaricaStringaRes(mnu.HelpContextID)
        End If
    Next
    pForm.PopupMenu Me.Controls(pstrNome), vbPopupMenuLeftAlign
    On Local Error GoTo 0
fine_mostraPopUp:
    MostraPopUpLingue = MidxMenu
    
    Unload Me
Exit Function

err_MostraPopUp:
    
    Call MXNU.MsgBoxEX(Err.Number & " - " & Err.Description, vbCritical, 1007)
    On Local Error GoTo 0
    
    Resume fine_mostraPopUp
End Function
Public Function MostraPopUpAllInOne(strValid As String) As String

'#If METODOXP = 1 Then
    Dim colConsoles As Object
    Dim cInteraction As Object
    Dim i As Integer
    
    If (Not MXALL Is Nothing) Then
        If (MXALL.Validation2ConsoleList(strValid, colConsoles)) Then
            MidxMenu = -1
            i = 0
            For Each cInteraction In colConsoles
                If i = 0 Then
                    AllInOneItem(i).Caption = cInteraction.ConsoleName & " - " & cInteraction.Description
                    i = i + 1
                Else
                    i = i + 1
                    Load AllInOneItem(i)
                    AllInOneItem(i).Caption = cInteraction.ConsoleName & " - " & cInteraction.Description
                End If
            Next
            If i > 0 Then
                If i = 1 Then
                    MostraPopUpAllInOne = Left(AllInOneItem(0).Caption, InStr(AllInOneItem(0).Caption, " - ") - 1)
                Else
                    MXNU.FrmMetodo.PopupMenu Me.Controls("AllInOne"), vbPopupMenuLeftAlign
                    If MidxMenu <> -1 Then
                        MostraPopUpAllInOne = Left(AllInOneItem(MidxMenu).Caption, InStr(AllInOneItem(MidxMenu).Caption, " - ") - 1)
                    End If
                    Unload Me
                End If
            End If
        End If
    End If
    Set colConsoles = Nothing
    Set cInteraction = Nothing
'#End If

End Function

Public Function MostraPopUpDesigner(pForm As Form, ByVal pstrNome As String) As Integer
 Dim mDocUser As MSXML2.DOMDocument
 Dim oNodeList As MSXML2.IXMLDOMNodeList
 Dim oUserNode As MSXML2.IXMLDOMNode
 Dim bolRes As Boolean
 Dim strActiveVersion As String
 Dim strVersion As String
 Dim i As Integer
 
    On Local Error GoTo err_MostraPopUp
    
    strActiveVersion = "Standard"
    ' Lettura del file xml "UsersSettings.xml" per la gestione delle versioni disponibili all'utente
    Set mDocUser = New MSXML2.DOMDocument
    bolRes = mDocUser.Load(MXNU.PercorsoDesigner & "\UsersSettings.xml")
    If (bolRes) Then
        Set oUserNode = mDocUser.selectSingleNode("settings/firm[@name='" & MXNU.DittaAttiva & "']/user[@name='" & LCase(MXNU.UtenteAttivo) & "']")
        If Not (oUserNode Is Nothing) Then
            strActiveVersion = oUserNode.Attributes.getNamedItem("activeversion").nodeValue
            Set oNodeList = mDocUser.selectNodes("settings/firm[@name='" & MXNU.DittaAttiva & "']/user[@name='" & LCase(MXNU.UtenteAttivo) & "']/version")
            If Not (oNodeList Is Nothing) Then
                i = 0
                For Each oUserNode In oNodeList
                    With oUserNode.Attributes
                        strVersion = .getNamedItem("name").nodeValue
                        If i <> 0 Then
                            ' La prima voce di menù è già presente (versione standard - cambio solo la caption)
                            Load DesignerItem(i)
                        End If
                        DesignerItem(i).Caption = strVersion
                        If strVersion = strActiveVersion Then
                            Me.Controls("DesignerItem")(i).Checked = True
                        Else
                            Me.Controls("DesignerItem")(i).Checked = False
                        End If
                        i = i + 1
                    End With
                Next oUserNode
                Set oNodeList = Nothing
            End If
        Else
            Me.Controls("DesignerItem")(0).Checked = True
        End If
        Set oUserNode = Nothing
        
    End If
    
    Me.Controls("Designer").Visible = True
    pForm.PopupMenu Me.Controls("Designer"), vbPopupMenuLeftAlign
    
    On Local Error GoTo 0

fine_mostraPopUp:
    Unload Me
Exit Function

err_MostraPopUp:
    
    Call MXNU.MsgBoxEX(Err.Number & " - " & Err.Description, vbCritical, 1007)
    On Local Error GoTo 0
    Resume fine_mostraPopUp

End Function

Private Sub AllInOneItem_Click(Index As Integer)
    MidxMenu = Index
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'Anomalia nr. 10385
    MidxMenu = Control.Id - 50001
    If MidxMenu > 13 Then
        mStrMostraCol = Control.Caption
        MidxMenu = 13
    End If
End Sub

Private Sub DesignerItem_Click(Index As Integer)
  Dim oDocUsers As MSXML2.DOMDocument
  Dim oUserNode As MSXML2.IXMLDOMNode
  Dim bolRes As Boolean
  
    ' Devo cambiare l'ActiveVersion del file xml e settare la VersioneAttiva del nucleo
    On Local Error GoTo ERR_DesignerItem
    Set oDocUsers = New MSXML2.DOMDocument
    bolRes = oDocUsers.Load(MXNU.PercorsoDesigner & "\UsersSettings.xml")
    If (bolRes) Then
        Set oUserNode = oDocUsers.selectSingleNode("settings/firm[@name='" & MXNU.DittaAttiva & "']/user[@name='" & LCase(MXNU.UtenteAttivo) & "']")
        If Not (oUserNode Is Nothing) Then
            If oUserNode.Attributes.getNamedItem("activeversion").nodeValue <> DesignerItem(Index).Caption Then
                oUserNode.Attributes.getNamedItem("activeversion").nodeValue = DesignerItem(Index).Caption
                bolRes = True
            Else
                bolRes = False
            End If
            
            ' Modifico la versione attiva nel nucleo
            MXNU.VersioneAttiva = DesignerItem(Index).Caption
                
            ' ... e ricarico eventualmente lo sfondo di versione
            Dim objSys As FileSystemObject
            Set objSys = New FileSystemObject
            If objSys.FileExists(MXNU.PercorsoPgm & "\images\sfondo." & MXNU.VersioneAttiva & ".jpg") Then
                metodo.Hide
                Set metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo." & MXNU.VersioneAttiva & ".jpg")
                metodo.Show
                DoEvents
            ElseIf objSys.FileExists(MXNU.PercorsoPgm & "\images\sfondo." & MXNU.VersioneAttiva & ".bmp") Then
                metodo.Hide
                Set metodo.Picture = LoadPicture(MXNU.PercorsoPgm & "\images\sfondo." & MXNU.VersioneAttiva & ".bmp")
                metodo.Show
                DoEvents
            End If
            Set objSys = Nothing

        Else
            MXNU.VersioneAttiva = "Standard"
        End If
        
        ' Visualizzo sulla barra di stato la versione attiva
        metodo.BarraStato.Panels("Designer").text = MXNU.VersioneAttiva
        ' Risetto la form attiva in modo da forzare l'ApplyDesign del componente Designer
        Set metodo.FormAttiva = metodo.ActiveForm
    
    End If
    If bolRes Then Call oDocUsers.Save(MXNU.PercorsoDesigner & "\UsersSettings.xml")
    On Local Error GoTo 0
   
EXIT_DesignerItem:
    Set oUserNode = Nothing
    Set oDocUsers = Nothing
    Exit Sub

ERR_DesignerItem:
  Dim lngErrCod As Long
  Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("ApplyDesign", lngErrCod, strErrDsc))
    Resume EXIT_DesignerItem
    Resume
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set frmPopUp = Nothing
End Sub

Private Sub GestDocRItem_Click(Index As Integer)
    MidxMenu = Index
End Sub

Private Sub GestDocRItemHide_Click(Index As Integer)
    mStrMostraCol = GestDocRItemHide(Index).Caption
End Sub

Private Sub GestDocTItem_Click(Index As Integer)
    MidxMenu = Index
End Sub

Private Sub LingueItem_Click(Index As Integer)
    Dim strLingua As String

    strLingua = Trim(LingueItem(Index).Caption)
    If MXNU.LinguaAttiva <> strLingua Then
        Call ChiudiFormAttive
        Call MXNU.CambiaRisorse(strLingua)
        
        metodo.MousePointer = vbHourglass
        #If ISMETODO2005 <> 1 Then
            Unload frmModuli
            Call CaricaModuliMetodo
        #Else
            Call CaricaModuliMetodo(True)
            Call CambiaRisorseCJ
        #End If
        frmModuli.ModuloAttivo = "Metodo98"
        Call AggiornaStatusBar
        metodo.MousePointer = vbNormal
    End If
    
End Sub

Private Sub mnuEcrItem_Click(Index As Integer)
    'frmMovimentiVBanco.mIntNumeroProtECR = Index
End Sub

'menu verticale

Public Function Reset()
    mnuVerticaleSel = ""
    indVerticaleSel = -1
End Function

Private Sub mnuNodiPrevCommCliItem_Click(Index As Integer)
     MidxMenu = Index
End Sub

Private Sub mnuSincStruttCCItem_Click(Index As Integer)
    MidxMenu = Index
End Sub

Private Sub MnuVertFigliItem_Click(Index As Integer)
    mnuVerticaleSel = "mnuVertFigliItem"
    indVerticaleSel = Index
End Sub

Private Sub MnuVertGruppiItem_Click(Index As Integer)
    mnuVerticaleSel = "mnuVertGruppiItem"
    indVerticaleSel = Index
End Sub

Private Sub mnuVertItem_Click(Index As Integer)
    mnuVerticaleSel = "mnuVertItem"
    indVerticaleSel = Index
End Sub

Private Sub mnuSSCCItem_Click(Index As Integer)
    MidxMenu = Index
End Sub

Private Sub mnuSSCCPalletItem_Click(Index As Integer)
    MidxMenu = Index
End Sub

Private Sub mnuSSCCProductItem_Click(Index As Integer)
    MidxMenu = Index
End Sub

