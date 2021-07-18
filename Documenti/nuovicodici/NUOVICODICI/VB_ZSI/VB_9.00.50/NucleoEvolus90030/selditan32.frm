VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelDitta 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   Caption         =   "Selezione Ditte"
   ClientHeight    =   4620
   ClientLeft      =   2295
   ClientTop       =   2610
   ClientWidth     =   7650
   ControlBox      =   0   'False
   FillColor       =   &H8000000A&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4620
   ScaleWidth      =   7650
   Begin VB.CommandButton com 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   540
      TabIndex        =   2
      Top             =   4080
      WhatsThisHelpID =   25007
      Width           =   1095
   End
   Begin VB.CommandButton com 
      Caption         =   "&Annulla"
      Height          =   375
      Index           =   1
      Left            =   6060
      TabIndex        =   1
      Top             =   4080
      WhatsThisHelpID =   25008
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListaDitte 
      Height          =   3555
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6271
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ImageDB"
      SmallIcons      =   "ImageDB"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nome Ditta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrizione"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fonte Dati"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tipo Server"
         Object.Width           =   8819
      EndProperty
   End
End
Attribute VB_Name = "frmSelDitta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

#If IsMetodo2005 Then
    'Per Metodo Evolus
    Public WithEvents mResize As MxResizer.ResizerEngine
Attribute mResize.VB_VarHelpID = -1
#End If

'====================================
'   dichiarazione costanti
'====================================
Const COL_NomeDitta = 0
Const COL_DscDitta = 1
Const COL_FonteDati = 2
Const COL_Driver = 3

Const COM_CONFERMA = 0
Const COM_ANNULLA = 1

'====================================
'   dichiarazione variabili
'====================================
Dim bolAnnullata As Boolean

Function CaricaListaDitte(strSelDitta As String) As Boolean
On Local Error Resume Next

Dim strAus As String
Dim intAus As Integer, cnt As Integer
Dim intCurD As Integer, intNumD As Integer
Dim vetDitte() As String
Dim vetCampi() As String
Dim objItem As ListItem
Dim bolInsert As Boolean
Dim strConnStr As String
Dim vetDitteAbilitate() As String

    CaricaListaDitte = False
    strAus = MXNU.LeggiProfilo(MXNU.PercorsoLocal & "\Ditte.ini", "DITTE", 0&, "")
    If (strAus <> "") Then
    
        vetDitteAbilitate = Split(MXNU.LeggiProfilo(MXNU.PercorsoLocal & "\Ditte.Ini", MXNU.UtenteSistema, "DitteAbilitate", ""), ",")
        
        CaricaListaDitte = True
        com(0).Enabled = False
        ListaDitte.ListItems.Clear
        ReDim vetDitte(0) As String
        intNumD = slice(strAus, Chr$(0), vetDitte())
        For intCurD = 0 To intNumD - 1
            strAus = MXNU.LeggiProfilo(MXNU.PercorsoLocal & "\Ditte.ini", "DITTE", vetDitte(intCurD), "")
            If (strAus <> "") Then
                ReDim vetCampi(2) As String
                intAus = slice(strAus, ";", vetCampi())
                If StrComp(MXNU.DittaAttiva, vetDitte(intCurD), vbTextCompare) <> 0 Then
                    bolInsert = True
                    If MXNU.MetodoXP Then
                        strConnStr = MXNU.GetstrConnection(vetCampi(0))
                        bolInsert = (InStr(1, strConnStr, "{Adaptive Server Anywhere 6.0}", vbTextCompare) = 0)
                    End If
                    If bolInsert Then
                        'Rif. Sviluppo Nr. 1749
                        If UBound(vetDitteAbilitate) >= 0 Then
                            bolInsert = (TrovaElementoVet(vetDitteAbilitate, vetDitte(intCurD)) >= 0)
                        Else
                            bolInsert = True
                        End If
                        If bolInsert Then
                            Set objItem = ListaDitte.ListItems.Add(, vetCampi(0), vetDitte(intCurD))
                            objItem.SubItems(COL_DscDitta) = vetCampi(2)
                            objItem.SubItems(COL_FonteDati) = vetCampi(0)
                            objItem.SubItems(COL_Driver) = vetCampi(1)
                        End If
                    End If
                    com(0).Enabled = True
                End If
            End If
        Next intCurD
        On Local Error Resume Next
        ListaDitte.SelectedItem = ListaDitte.ListItems(strSelDitta)
        If (Err <> 0) Then ListaDitte.SelectedItem = ListaDitte.ListItems(1)
        On Local Error GoTo 0
    End If
End Function

Sub DefLingua()
    Me.Caption = MXNU.CaricaStringaRes(23026)
    Call MXNU.LeggiRisorseControlli(Me)
    'Caricamento intestazioni colonne Anomalia 2621
    ListaDitte.ColumnHeaders(1).Text = MXNU.CaricaStringaRes(31423)
    ListaDitte.ColumnHeaders(2).Text = MXNU.CaricaStringaRes(30005)
    ListaDitte.ColumnHeaders(3).Text = MXNU.CaricaStringaRes(31424)
    ListaDitte.ColumnHeaders(4).Text = MXNU.CaricaStringaRes(31425)
End Sub

Public Function SelezioneDitta(strDitta As String) As Boolean
    SelezioneDitta = False
    If (CaricaListaDitte(strDitta)) Then
        SelezioneDitta = False
        Me.Show vbModal
        'risultato selezione

        If bolAnnullata Then
            strDitta = ""
        Else
            'inizio rif.sch. A4290
            If Not (ListaDitte.SelectedItem Is Nothing) Then
                SelezioneDitta = True
                strDitta = ListaDitte.SelectedItem.Text
            End If
            'fine rif.sch. A4290
        End If
    End If
End Function

Private Sub com_Click(Index As Integer)
    Select Case Index
        Case COM_CONFERMA
            bolAnnullata = False
        Case COM_ANNULLA
            bolAnnullata = True
    End Select
    Me.Hide
    DoEvents
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then Call com_Click(COM_ANNULLA)
End Sub

Private Sub Form_Load()

    Me.MousePointer = vbHourglass
    Call DefLingua
    Call CentraFinestra(Me.hwnd)
    Me.MousePointer = vbDefault
#If IsMetodo2005 Then
    'Inzializzazione Form per Metodo Evolus
    Call CambiaColoriControlli(Me)   'Da Mettere PRIMA di mResize
    On Local Error Resume Next
    Set mResize = New MxResizer.ResizerEngine
    If (Not mResize Is Nothing) Then
            Call mResize.Initialize(Me, , , , , 0, True, metodo, MXNU)
    End If
#End If
Call CentraFinestra(Me.hwnd)
Call CambiaCharSet(Me)
On Local Error GoTo 0
End Sub

Private Sub Form_Terminate()
    Set frmSelDitta = Nothing
End Sub

Private Sub ListaDitte_DblClick()
    Call com_Click(COM_CONFERMA)
End Sub

Private Sub ListaDitte_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then Call com_Click(COM_CONFERMA)
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
