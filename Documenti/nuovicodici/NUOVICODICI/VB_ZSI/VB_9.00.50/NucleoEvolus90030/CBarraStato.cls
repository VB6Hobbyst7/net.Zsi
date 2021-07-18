VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBarraStato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum setStatusBarPanelID
    STATUSBAR_ID_PANELINFO = 100   'Cambiato da 0 a 100 per Anomalia 10797 (sul pannello di stato con ID=0 viene messo il testo della proprietà DescriptionText degli elementi della CommandBar Codejock)
    STATUSBAR_ID_PANELDITTAANNO = 101
    STATUSBAR_ID_PANELUTENTE = 102
    STATUSBAR_ID_PANELSAVE = 103
    STATUSBAR_ID_PANELDESIGNER = 104
    STATUSBAR_ID_PANELLINGUA = 105
End Enum


Private colKeyPanelToID As New Collection

Public Property Get Panels(PanelID As Variant) As XtremeCommandBars.StatusBarPane
    
    If IsNumeric(PanelID) Then
        Set Panels = metodo.CommandBars.StatusBar.FindPane((PanelID))
    Else
        Set Panels = metodo.CommandBars.StatusBar.FindPane(colKeyPanelToID(PanelID))
    End If
End Property


Private Sub Class_Initialize()
    colKeyPanelToID.Add STATUSBAR_ID_PANELINFO, "info"
    colKeyPanelToID.Add STATUSBAR_ID_PANELDITTAANNO, "dittaanno"
    colKeyPanelToID.Add STATUSBAR_ID_PANELUTENTE, "utente"
    colKeyPanelToID.Add STATUSBAR_ID_PANELSAVE, "save"
    colKeyPanelToID.Add STATUSBAR_ID_PANELDESIGNER, "designer"
    colKeyPanelToID.Add STATUSBAR_ID_PANELLINGUA, "lingua"
End Sub


Private Sub Class_Terminate()
    Set colKeyPanelToID = Nothing
End Sub


