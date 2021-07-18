VERSION 5.00
Object = "{E3D4F235-2CDC-48E5-B30A-E85333A88160}#1.0#0"; "mxctrl.ocx"
Begin VB.Form frmExtChild 
   ClientHeight    =   2715
   ClientLeft      =   2070
   ClientTop       =   3885
   ClientWidth     =   5940
   Icon            =   "frmExtChild.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   5940
   Begin MXCtrl.MWEtichetta lblerr 
      Height          =   975
      Left            =   120
      Top             =   420
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      Caption         =   "Impossibile caricare l'estensione"
   End
End
Attribute VB_Name = "frmExtChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents MWAgt1 As MXKit.CAgenteAuto
Attribute MWAgt1.VB_VarHelpID = -1
Public FormProp As New CFormProp
Public NomeEstensione  As String
Public FunzioniM98 As CFunzioniMetodo98
Public NomeWrapper As String
Public visibile As Boolean
Public objAnagrEXTRA As Object
Private MctlExt As VBControlExtender
Private MctlWrapper As VBControlExtender
Public ExtScaricata As Boolean

#If ISMETODO2005 Then
    'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1
#End If

Private Sub Form_Load()
    Dim intRes As Integer
    Dim colAmb As Collection
    Dim colObj As Collection
    Dim strNomeExt As String
    Dim intq As Integer
    On Local Error GoTo Err_Load
    
    Me.HelpContextID = FormProp.FormID
    
    #If ISM98SERVER <> 1 Then
    metodo.MousePointer = vbHourglass
    #End If
    
    Set colObj = New Collection
    Set FunzioniM98 = New CFunzioniMetodo98
    'inizializzazione agenti
    Set MWAgt1 = MXAA.CreaCAgenteAuto()
    intRes = MXAA.RegistraEventiFrm(Me, MWAgt1)
    
    'impostazione accessi
    On Local Error Resume Next
    'oggetti std di M98
    Licenses.Add NomeEstensione
    On Local Error GoTo Err_Load
    If NomeWrapper = "" Then
    
        Set MctlExt = Me.Controls.Add(NomeEstensione, OGGETTO_ESTENSIONE, Me)
    Else
        On Local Error Resume Next
        Licenses.Add NomeWrapper
        On Local Error GoTo Err_Load
        Set MctlWrapper = Me.Controls.Add(NomeWrapper, OGGETTO_WRAPPER_ESTENSIONE, Me)
        Set MctlExt = MctlWrapper.object.CaricaEstensione(NomeEstensione)
        MctlWrapper.Visible = True
    End If
    
    Set colAmb = Ambienti2Collection(MXNU.CheckOwner(MctlWrapper))
    colObj.Add hndDBArchivi
    
    If (objAnagrEXTRA Is Nothing) Then
        Call MctlExt.object.Inizializza(Me, colAmb, colObj, Nothing)
    Else
        Call MctlExt.object.Inizializza(Me, colAmb, colObj, objAnagrEXTRA)
    End If
    
fine_Load:
    Set colAmb = Nothing
    Set colObj = Nothing
    
    ' Rif. scheda #9022 (aggiunto test su MctlExt)
    If Not ExtScaricata And Not (MctlExt Is Nothing) Then 'Altrimenti su Evolus và in loop e ricarica l'estensione (vedi UnloadFormExt)
        MctlExt.TabIndex = 0
        MctlExt.Visible = True
    
        'dimensionamento
        Me.Height = MctlExt.Height + 390
        Me.width = MctlExt.width + 90
    
        'mostra la finestra
        Call CentraFinestra(Me.hwnd)
        If MXNU.MetodoXP And Not (MctlExt Is Nothing) Then
            'Call ModificaLayoutControlli(Me)
            intq = InStr(NomeEstensione, ".")
            If intq > 1 Then
                strNomeExt = Left(NomeEstensione, intq - 1)
            End If
            Select Case LCase(strNomeExt)
                Case "analisibilancio", "exttargetagenti", "extcommcli", "e98comm", "extmps", "extmetodo", "mxexport", "extspedizionieri", "exteuris"
                Case Else
                    If Not (MctlExt.object Is Nothing) Then
                        Call ModificaLayoutControlli(MctlExt.object)
                    End If
            End Select
        End If
    
        'Disabilito il bottone di Designer nel caso di estensione (finché non verranno gestite in MYERP)
        #If ISM98SERVER <> 1 Then
        metodo.Barra.Buttons(idxBottoneDesigner).Enabled = False
        #End If
        
        
        If Not visibile Then
            If Not MctlExt Is Nothing Then Call CambiaCharSet(MctlExt.object)
            'Per Metodo Evolus
        On Local Error Resume Next
            'Call CambiaColoriControlli(Me)
            Dim bolAbilitaRezizer As Boolean
            Me.Left = -20000
            Me.Show
            DoEvents
            Call CambiaColoriControlli(MctlExt.object)
            DoEvents
        #If ISMETODO2005 Then
            Select Case UCase(Left(NomeEstensione, InStr(NomeEstensione, ".") - 1))
                Case Else: bolAbilitaRezizer = True
            End Select
            If bolAbilitaRezizer Then
                Set mResize = New MxResizer.ResizerEngine
                If (Not mResize Is Nothing) Then
                    Call mResize.Initialize(Me, , , , , 0, True, metodo, MXNU)
                End If
            End If
        #End If
            Call CentraFinestra(Me.hwnd)
            DoEvents
            If TemaConGradiente Then
                Me.BackColor = vbInactiveTitleBarText
            Else
                Me.BackColor = vbButtonFace
            End If
            'Altrimenti se ci sono linguette non viene abilitata la prima scheda
            MctlExt.object.Controls("Ling")(0).SetFocus
            On Local Error GoTo 0
            Me.Show
        End If
    End If
    
    #If ISM98SERVER <> 1 Then
    metodo.MousePointer = vbDefault
    #End If

Exit Sub

Err_Load:
    Dim coderr&, dscerr$
    coderr = Err.Number
    dscerr = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("Caricamento Estensioni", coderr, dscerr))
    lblerr.Visible = True
    Resume fine_Load
    Resume
    
    ' If Err.Number = 713 Then
         'Err.Clear
         'On Local Error GoTo fine_Load:
         'Dim Nu98 As Object
         'Set Nu98 = CreateObject("EXtBus.ClassLoader")
         'Dim Nu98 As Nucleo98.Loader
         'Set Nu98 = New Nucleo98.Loader
         'Set objSch.Parent.FunzioniM98 = New CFunzioniMetodo98
         'Set colAmb = Ambienti2Collection(MXNU.CheckOwner(MctlWrapper))
         'colObj.Add hndDBArchivi
         'Nu98.PNomeEstensione = NomeEstensione
         'Nu98.PNomeWrapper = NomeWrapper
         'Nu98.PcolAmb = colAmb
         'Nu98.PcolObj = colObj
         'Nu98.PObjAnagr = Nothing
         'Call Nu98.LoadEXT
         'Set Nu98 = Nothing
     'End If
     '*****************
Exit Sub
    
    
End Sub


Public Function AzioniMetodo(setAzione As enmFunzioniMetodo, Optional varparametro As Variant) As Variant

    Select Case setAzione
'        Case MetFVisUtenteModifica
'            On Local Error Resume Next
'            Select Case DammiValidazione(MXNU.NomeControllo(ctlExt.ActiveControl))
'                Case artCommerciali
'                    Call MXCT.VisDatiUtenteModifica(FrmVisUtMod, xArtComm.NOMETABELLA, "CodiceArt='" & txtb(0).Text & "' AND Esercizio=" & Trim(txtEse(0).Text), txtb(0).Text & "-" & txtEse(0).Text, "", Nothing)
'                Case artProduzione
'                    Call MXCT.VisDatiUtenteModifica(FrmVisUtMod, xArtProd.NOMETABELLA, "CodiceArt='" & txtb(0).Text & "' AND Esercizio=" & Trim(txtEse(1).Text), txtb(0).Text & "-" & txtEse(1).Text, "", Nothing)
'                Case Else
'                    Call MXCT.VisDatiUtenteModifica(FrmVisUtMod, xArt.NOMETABELLA, "Codice='" & txtb(0).Text & "'", txtb(0).Text & "-" & txtb(2).Text, "MAG", SSExtra)
'            End Select
'            On Local Error GoTo 0
    Case Else
            AzioniMetodo = MctlExt.object.AzioniMetodo(setAzione, varparametro)
    End Select
End Function

Public Function ContrEXT(Optional Index As Variant, Optional Index2 As Variant) As Object
    
        If IsMissing(Index) Then
            
            Set ContrEXT = MctlExt.object.Controls
        
        ElseIf IsMissing(Index2) Then
            
            Set ContrEXT = MctlExt.object.Controls(Index)
        
        Else
            
            Set ContrEXT = MctlExt.object.Controls(Index)(Index2)
        
        End If

End Function

Public Sub EseguiComandoBatch(colParametri As Variant, ByRef filedilog As Variant)
    #If ISM98SERVER = 1 Then
        Load MDIComStd
        DoEvents
        Load Me
        MDIComStd.Visible = False
        
        'implementare "ChiamaFunzioneInterna" x un metodo definito come std per lanciare
        'operazioni dallo schedulatore!
        Dim vntArgs(1) As Variant
        Set vntArgs(0) = colParametri
        vntArgs(1) = filedilog
        Call ChiamaFunzioneInterna("EseguiComandoBatch", vntArgs)
        filedilog = vntArgs(1)
    #End If
End Sub

Private Sub Form_Paint()
    'Call SchedaOmbreggiaControlli(Me)
    If Me.Controls.Count > 1 Then Call SchedaOmbreggiaControlli(Me.Controls(1))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

#If ISMETODO2005 Then
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

Private Sub Form_Unload(Cancel As Integer)
    
    On Local Error Resume Next
    'RIF.A#6646 - lo scaricamento dell'agente con Termina deve essere fatto prima della distruzione
    'dell'estensione altrimenti provoca problemi con lo scaricamento della collection colSubclassedObject
    If (Not MWAgt1 Is Nothing) Then
        Call MWAgt1.Termina
        Set MWAgt1 = Nothing
    End If
    'DoEvents
    If Not MctlExt Is Nothing Then
        MctlExt.Visible = False
        'DoEvents
        If Not MctlExt.object Is Nothing Then MctlExt.object.Termina
        If MctlWrapper Is Nothing Then
            ' Anomalia zoom n.ro 6084 (invertito l'ordine delle 2 righe che seguono)
            Me.Controls.Remove OGGETTO_ESTENSIONE
            Set MctlExt = Nothing
        Else
            MctlWrapper.Visible = False
            Call MctlWrapper.object.ScaricaEstensione
            'Set MctlExt = Nothing
            Me.Controls.Remove OGGETTO_WRAPPER_ESTENSIONE
            'Set MctlWrapper = Nothing
        End If
    End If
    Set FunzioniM98 = Nothing
    Set MctlExt = Nothing
    Set MctlWrapper = Nothing
    'Set frmExtChild = Nothing

    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
    On Local Error GoTo 0
    Set FormProp = Nothing
    DoEvents
'    Set frmExtChild = Nothing
    
    #If ISM98SERVER <> 1 Then
    metodo.Barra.Buttons(idxBottoneDesigner).Enabled = True
    #End If
        
End Sub

Public Function ChiamaFunzioneInterna(ByVal NomeFunzione As String, Args() As Variant) As Boolean
    Select Case UCase(NomeFunzione)
        Case "ESEGUICOMANDOBATCH"
            CallByName MctlExt.object, NomeFunzione, VbMethod, Args(0), Args(1)
        Case Else
            CallByName MctlExt.object, NomeFunzione, VbMethod, Args(0)
    End Select
End Function

#If ISMETODO2005 Then
    Private Sub mResize_AfterInitialize()
        If Not (MctlExt Is Nothing) And Not (mResize Is Nothing) Then
            On Local Error Resume Next
            Call MctlExt.object.ResizeAfterInitialize(mResize)
            On Local Error GoTo 0
        End If
    End Sub

    'Per Metodo Evolus
    Private Sub mResize_AfterResize()
        If Not (MctlExt Is Nothing) Then
           Call AvvicinaLing(MctlExt.object)
           On Local Error Resume Next
           If Not (mResize Is Nothing) Then Call MctlExt.object.ResizeAfterResize(mResize)
           On Local Error GoTo 0
        End If
    End Sub
    
    Public Sub InitToolbarExt()
        Call InitToolbarForm(MctlExt.object)
    End Sub
    
    Public Sub GestioneToolbarExt(btnId As Long)
        Call GestioneToolBut2005(btnId)
    End Sub
    
#End If
