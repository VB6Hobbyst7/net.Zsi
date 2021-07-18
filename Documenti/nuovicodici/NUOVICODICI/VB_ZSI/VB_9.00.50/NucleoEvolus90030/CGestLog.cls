VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGestLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MstrTitolo As String

Private McolLogHandles As New Collection
Private McolLogPanesID As New Collection
Private NLogPanels As Long
Private MLastPaneID As Long
Private McolAddedMessages As Collection

Private MLngIndiceNodo As Long

'Per Anomalia 11649 e 11978
Public mvarUseExistingLogPanel As Boolean
Public HideFirstLog As Boolean
Public Sub AddMessageToLog(ByVal strLabelText As String, _
    Optional ByVal enmTipoMsg As setTipoMessaggio = msgNessuno, _
    Optional ByVal vntKey As Variant, _
    Optional ByVal objRelativeTo As MSComctlLib.Node = Nothing)
    
    Dim strMsg As String
    
    If McolAddedMessages Is Nothing Then
        Set McolAddedMessages = New Collection
        MLngIndiceNodo = 0
    End If
    
    If IsMissing(vntKey) Then
        vntKey = "_" & MLngIndiceNodo
        MLngIndiceNodo = MLngIndiceNodo + 1
    End If
    strMsg = strLabelText & "|"
    strMsg = strMsg & enmTipoMsg & "|"
    strMsg = strMsg & vntKey & "|"
    If Not (objRelativeTo Is Nothing) Then
        strMsg = strMsg & objRelativeTo.key
    Else
        strMsg = strMsg & " "
    End If
    McolAddedMessages.Add strMsg
    
End Sub

Public Property Get HasMessages() As Boolean
    If Not (McolAddedMessages Is Nothing) Then
        HasMessages = McolAddedMessages.Count > 0
    Else
        HasMessages = False
    End If
End Property


Public Function IsLogPane(ByVal PaneID As Long) As Boolean
    Dim vntPanel As Variant
    IsLogPane = (EsisteElementoCollection(McolLogPanesID, CStr(PaneID)))
End Function

'같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같�
'           FUNZIONI PUBBLICHE DELLA FORM
'같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같같�
Public Sub MostraFileLog(ByVal strNomeFile As String, _
    Optional lngResMsgConferma As Long = 0, _
    Optional lngResNodoRoot As Long = 23117, _
    Optional bolMostraModale As Boolean = False)
Dim lngLOF As Long
Dim bolApri As Boolean
Dim LogPaneID As Long
Dim PnlLog As XtremeDockingPane.Pane


    On Local Error Resume Next
    lngLOF = FileLen(strNomeFile)
    If (Err = 0) And (lngLOF > 0) Then
    'rif sviluppo 338
        If (lngLOF < 50000) Then
            '0.5MB -> apro log
            If lngResMsgConferma = 0 Then
                bolApri = True
            Else
                bolApri = MXNU.MsgBoxEX(lngResMsgConferma, vbYesNo + vbQuestion, "") = vbYes
            End If
            If bolApri Then
                If lngResNodoRoot = 0 Then lngResNodoRoot = 23117
                Dim frmLogContainer As New frmLogPanel
                With frmLogContainer
                    .objLog.pLogTitle = MXNU.CaricaStringaRes(lngResNodoRoot)
                    If .objLog.ImportLogFile(strNomeFile) Then
                        Call .objLog.AttivaNodo(AttNodoPrimo)
                        If NLogPanels = 0 Then
                            LogPaneID = ID_PANE_TASKBAR1
                            'Set PnlLog = metodo.DockingPaneManager.CreatePane(LogPaneID, 250, 200, DockRightOf, metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR2))
                            Set PnlLog = metodo.DockingPaneManager.CreatePane(LogPaneID, 250, 200, DockBottomOf, Nothing)
                            'Set PnlLog = metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1)
                            PnlLog.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "VisLog")
                            PnlLog.Hide
                            If Not (metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR2) Is Nothing) Then
                                'Anomalia 9065
                                If Not metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR2).Closed Then
                                    PnlLog.AttachTo metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR2)
                                End If
                            End If
                            NLogPanels = NLogPanels + 1
                        Else
                            'Aggiungo un nuovo pannello
                            LogPaneID = BOTTOM_PANELS_OFFSET + NLogPanels
                            Set PnlLog = metodo.DockingPaneManager.CreatePane(LogPaneID, 250, 200, DockBottomOf, Nothing)
                            PnlLog.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "VisLog")
                            PnlLog.Hide
                            PnlLog.AttachTo metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1)
                            NLogPanels = NLogPanels + 1
                        End If
                        If MstrTitolo = "" Then
                            'metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1).Title = MXNU.CaricaStringaRes(23117)
                            'PnlLog.Title = MXNU.CaricaStringaRes(23117) & " - " & Mid(strNomeFile, InStrRev(strNomeFile, "\") + 1)
                            PnlLog.Title = "Log - " & Mid(strNomeFile, InStrRev(strNomeFile, "\") + 1)
                            'If NLogPanels > 0 Then PnlLog.Title = PnlLog.Title & " " & NLogPanels
                        Else
                            'metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1).Title = MstrTitolo
                            PnlLog.Title = MstrTitolo
                        End If
                        .Caption = PnlLog.Title
                        'metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1).handle = Me.hwnd
                        'metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1).Select
                        PnlLog.Handle = .hwnd
                        McolLogHandles.Add frmLogContainer.hwnd, CStr(PnlLog.ID)
                        McolLogPanesID.Add PnlLog.ID, CStr(PnlLog.ID)
                        metodo.DockingPaneManager.RecalcLayout
                        PnlLog.Selected = False
                        'PnlLog.Selected = True
                        PnlLog.Select
                        'If bolMostraModale Then Call Me.Show(vbModal) Else Call Me.Show(vbModeless)
                    End If
                End With
            End If
        Else
            '>0.5MB -> apro applicazione associata
            Dim strPath As String
            Dim strFile As String
            If scomponiPathFile(strNomeFile, strPath, strFile) Then
                Call EseguiAppAssociata(strPath, strFile)
            End If
        End If
    End If
    On Local Error GoTo 0
End Sub


Public Sub ScaricaLog(ByVal PaneID As Long)
    If IsLogPane(PaneID) Then
        Dim i As Long
        For i = 0 To Forms.Count - 1
            Dim hwnd As Long
            hwnd = Forms(i).hwnd
            If hwnd = McolLogHandles(CStr(PaneID)) Then
                Unload Forms(i)
                'If bolChiudiPane Then
                '    If PaneID <> ID_PANE_TASKBAR1 Then
                '        metodo.DockingPaneManager.FindPane(PaneID).Close
                '    End If
                'End If
                McolLogHandles.Remove CStr(PaneID)
                McolLogPanesID.Remove CStr(PaneID)
                If NLogPanels > 0 Then
                    NLogPanels = NLogPanels - 1
                End If
                Exit For
            End If
        Next i
    End If
            
End Sub


Public Sub ShowAddedMessages()
    If Not Me.HasMessages Then Exit Sub
    
    Dim frmL As frmLogPanel
    Dim vntMsg As Variant
    Dim objLog As Node
    Dim objNodeRelative As MSComctlLib.Node
    Dim vetMsg() As String
    Dim LogPaneID As Long
    Dim PnlLog As XtremeDockingPane.Pane
    Dim bolCreatePane As Boolean
    
    If NLogPanels = 0 Then
        bolCreatePane = True
    Else
        bolCreatePane = ((Not mvarUseExistingLogPanel) Or Not EsisteElementoCollection(McolLogHandles, CStr(MLastPaneID)))
    End If
    If bolCreatePane Then
        Set frmL = New frmLogPanel
        For Each vntMsg In McolAddedMessages
            vetMsg = Split(vntMsg, "|")
            Call frmL.objLog.AddMessageToLog(vetMsg(0), vetMsg(1), vetMsg(2))
        Next
        With frmL
            If NLogPanels = 0 Then
                LogPaneID = ID_PANE_TASKBAR1
            Else
                LogPaneID = BOTTOM_PANELS_OFFSET + NLogPanels
            End If
            Set PnlLog = metodo.DockingPaneManager.CreatePane(LogPaneID, 250, 200, DockBottomOf, Nothing)
            PnlLog.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "VisLog")
            PnlLog.Hide
            If Not (metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR2) Is Nothing) Then
                'Anomalia 9065
                If Not metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR2).Closed Then
                    PnlLog.AttachTo metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR2)
                End If
            End If
            NLogPanels = NLogPanels + 1
            If MstrTitolo = "" Then
                'metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1).Title = MXNU.CaricaStringaRes(23117)
                PnlLog.Title = MXNU.CaricaStringaRes(23117)
                'If NLogPanels > 0 Then PnlLog.Title = PnlLog.Title & " " & NLogPanels
            Else
                'metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1).Title = MstrTitolo
                PnlLog.Title = MstrTitolo
            End If
            .Caption = PnlLog.Title
            'metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1).handle = Me.hwnd
            'metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1).Select
            PnlLog.Handle = .hwnd
            McolLogHandles.Add .hwnd, CStr(PnlLog.ID)
            McolLogPanesID.Add PnlLog.ID, CStr(PnlLog.ID)
            If Not HideFirstLog Then   'Anomalia 11649: per evitare lo sfarfallio iniziale se si stanno generando log in sequenza; vedi Stampa Differita Doc.
                metodo.DockingPaneManager.RecalcLayout
                PnlLog.Selected = False
                PnlLog.Selected = True
            End If
            MLastPaneID = PnlLog.ID
        End With
    Else
        Dim i As Long
        'Cerco la form associata all'ultimo pannello creato
        If EsisteElementoCollection(McolLogHandles, CStr(MLastPaneID)) Then
            For i = 0 To Forms.Count - 1
                Dim hwnd As Long
                hwnd = Forms(i).hwnd
                If hwnd = McolLogHandles(CStr(MLastPaneID)) Then
                    Set frmL = Forms(i)
                    Exit For
                End If
            Next i
            If Not (frmL Is Nothing) Then
                For Each vntMsg In McolAddedMessages
                    vetMsg = Split(vntMsg, "|")
                    Call frmL.objLog.AddMessageToLog(vetMsg(0), vetMsg(1), vetMsg(2))
                Next
                '***********************************************************************
                'Anomalia 11649:
                'Non visualizzo il log mentre li sto generando, altrimenti continua a sfarfallare il pannello e si rallenta la procedura
                'Richiamare il metodo ShowLastLog alla fine dell'elaborazione (vedi Stampa Differita frmFiltroComp)
                '***********************************************************************
                'Set PnlLog = metodo.DockingPaneManager.FindPane(MLastPaneID)
                'If Not (PnlLog Is Nothing) Then
                '    metodo.DockingPaneManager.RecalcLayout
                '    PnlLog.Selected = False
                '    PnlLog.Selected = True
                'End If
            End If
        End If
    End If
End Sub


Public Sub ShowLastLog()
    Dim PnlLog As XtremeDockingPane.Pane
    Set PnlLog = metodo.DockingPaneManager.FindPane(MLastPaneID)
    If Not (PnlLog Is Nothing) And MLastPaneID <> 0 Then
        metodo.DockingPaneManager.RecalcLayout
        PnlLog.Selected = False
        PnlLog.Selected = True
    Else
        If MLastPaneID <> 0 Then HideFirstLog = False
        Call ShowAddedMessages
        Set PnlLog = metodo.DockingPaneManager.FindPane(MLastPaneID)
        If Not (PnlLog Is Nothing) Then
            PnlLog.Selected = False
            PnlLog.Selected = True
        End If
    End If
    HideFirstLog = False
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    Dim PnlID As Long
    'Chiudo tutti i panes relativi ai logs aperti
    Do While McolLogPanesID.Count > 0
        'If McolLogPanesID(McolLogPanesID.Count) <> ID_PANE_TASKBAR1 Then
            PnlID = McolLogPanesID(McolLogPanesID.Count)
            ScaricaLog McolLogPanesID(McolLogPanesID.Count)
            DoEvents
            Dim p As XtremeDockingPane.Pane
            Set p = metodo.DockingPaneManager.FindPane(PnlID)
            If Not p Is Nothing Then
                metodo.DockingPaneManager.DestroyPane p
            End If
            
            
            'metodo.DockingPaneManager.FindPane(PnlID).Close
            'DoEvents
            'metodo.DockingPaneManager.RecalcLayout
        'Else
        '    'Per il pannello aggiunto in avvio di Metodo chiudo solo la form contenuta
        '    Call ScaricaLog(McolLogPanesID(McolLogPanesID.Count))
        '    metodo.DockingPaneManager.FindPane(ID_PANE_TASKBAR1).Title = "Log"
        'End If
    Loop
    metodo.DockingPaneManager.RecalcLayout
    If Not (McolAddedMessages Is Nothing) Then Set McolAddedMessages = Nothing
End Sub

Public Sub ClearLog()
    If Me.HasMessages Then
        Set McolAddedMessages = Nothing
        Set McolAddedMessages = New Collection
        MLngIndiceNodo = 0
    End If
End Sub

Public Property Get UseExistingLogPanel() As Boolean
    UseExistingLogPanel = mvarUseExistingLogPanel
End Property

Public Property Let UseExistingLogPanel(ByVal newValue As Boolean)
    mvarUseExistingLogPanel = newValue
    'Azzero l'ID dell'ultimo pannello in modo che ne venga creato uno nuovo (vedi funzione ShowLastLog)
    If Not newValue Then MLastPaneID = 0
End Property


