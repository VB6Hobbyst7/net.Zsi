VERSION 5.00
Object = "{A71FFC9D-B854-4499-A56D-5D144C9D3DD2}#1.0#0"; "mxctrl.ocx"
Begin VB.Form FrmTrovaGen 
   Caption         =   "Trova"
   ClientHeight    =   1815
   ClientLeft      =   3975
   ClientTop       =   5115
   ClientWidth     =   5745
   Icon            =   "Trovagen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5745
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Trova ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      HelpContextID   =   23350
      Left            =   60
      TabIndex        =   5
      Top             =   60
      WhatsThisHelpID =   23350
      Width           =   4125
      Begin MXCtrl.MWSchedaBox MWSchedaBox1 
         Height          =   435
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   767
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
         ScaleWidth      =   3975
         ScaleHeight     =   435
         FillWithGradient=   0   'False
         Begin VB.TextBox txtb 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   60
            Width           =   3735
         End
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Ricerca per ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      HelpContextID   =   11192
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   4155
      Begin VB.ComboBox cmb 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3735
      End
   End
   Begin VB.CommandButton Com 
      Caption         =   "&Fine"
      Height          =   345
      HelpContextID   =   25006
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   1080
      WhatsThisHelpID =   25006
      Width           =   1125
   End
   Begin VB.CommandButton Com 
      Caption         =   "&Trova"
      Height          =   345
      HelpContextID   =   23350
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Top             =   540
      WhatsThisHelpID =   25025
      Width           =   1125
   End
End
Attribute VB_Name = "FrmTrovaGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
DefInt A-Z

'Per Metodo Evolus
Private mResize As Object


'parametri della funzione
Public mySS As FPSpreadADO.fpSpread
Public ColTrv As Collection
Public ColOrd As Long
'costanti bottoni
Const COM_SUC = 0
Const COM_ANN = 1

Dim Succ As Long
Dim prevOpMode As MXSpread.ssOperationMode

Sub com_Click(Index As Integer)
    Select Case Index
        Case COM_SUC
            Call Go_Ricerca
        Case COM_ANN
            Unload Me
    End Select
End Sub

Sub End_Ricerca()
    mySS.OperationMode = prevOpMode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        Call Unload(Me)
    ElseIf (KeyAscii = vbKeyReturn) Then
       Call com_Click(COM_SUC)
    End If
End Sub

Sub Form_Load()
    Call Trova_Imposta
    Succ = 1&
    Call MXNU.LeggiRisorseControlli(Me)
    Me.Caption = MXNU.CaricaStringaRes(23350)
    If MXNU.ISMETODO2005 Then   'Non uso la variabile di compilazione condizionale in quanto questa form è condivisa con il Business
        Call CambiaColoriControlli(Me)
        On Local Error Resume Next
        Set mResize = CreateObject("MxResizer.ResizerEngine")
        If (Not mResize Is Nothing) Then
            Call mResize.Initialize(Me, , , , , 0, True, MXNU.FrmMetodo, MXNU)
        End If
        On Local Error GoTo 0
    End If
    Call CentraFinestra(Me.hwnd)
    'Rif. sviluppo #2234
    Call CambiaCharSet(Me)

End Sub


Sub Form_Unload(Cancel As Integer)
    Call End_Ricerca

    Set mySS = Nothing
    Set ColTrv = Nothing
    Set FrmTrovaGen = Nothing
End Sub

Function Go_Ricerca() As Integer
    Dim q As Integer
    Dim valRic As Variant
    Dim ColRic As Long
    Dim ColType As Integer
    Dim okRic As Integer
    Dim intMsg As Integer
    Dim Row As Long, Col As Long
    Dim Valore As Variant

    On Local Error GoTo err_Ricerca
    'imposta variabili
    valRic = Replace(txtb(0).text, "#", "!")    'Anomalia nr. 9430
    ColRic& = cmb.listIndex
    If ColRic = 0 Then
        For Row& = Succ& To mySS.DataRowCnt
            For Col& = 1 To mySS.MaxCols
                q% = mySS.GetText(Col&, Row&, Valore)
                okRic% = (RTrim(Replace(Valore, "#", "!")) Like valRic)     'Anomalia nr. 9430
                If (okRic%) Then
                    mySS.ReDraw = False
                    mySS.OperationMode = SS_OP_MODE_NORMAL
                    mySS.Col = Col&
                    mySS.Row = Row&
                    If mySS.Lock Then    'Anomalia nr. 10795
                        mySS.Col = ssColGetFirstActivable(mySS, Row)
                    End If
                    ssCellActive mySS, mySS.Col, mySS.Row, True     'Anomalia nr. 10073
                    mySS.OperationMode = SS_OP_MODE_SINGLE_SELECT
                    mySS.ReDraw = True
                    okRic% = True
                    Succ& = Row& + 1
                    GoTo fine_Ricerca
                End If
            Next Col
        Next Row
    Else
        ColType% = cmb.ItemData(cmb.listIndex)
        okRic% = True: intMsg = 0
        'validazione campo
        Select Case ColType%
            Case SS_CELL_TYPE_DATE
                If (Not IsDate(Format$(valRic, MXNU.Formato_Data))) Then
                    okRic% = False
                    intMsg = 2125
                    GoTo fine_Ricerca
                End If
            Case SS_CELL_TYPE_EDIT, SS_CELL_TYPE_STATIC_TEXT, SS_CELL_TYPE_COMBOBOX
            Case SS_CELL_TYPE_FLOAT, SS_CELL_TYPE_INTEGER
                If (Not IsNumeric(valRic)) Then
                    okRic% = False
                    intMsg = 2126
                    GoTo fine_Ricerca
                End If
        End Select
        okRic% = False
        For Row& = Succ& To mySS.DataRowCnt
            q% = mySS.GetText(ColTrv(ColRic&), Row&, Valore)
            'confronto valore letto
            Select Case ColType%
                Case SS_CELL_TYPE_DATE
                    If (IsDate(Format$(Valore, MXNU.Formato_Data))) Then
                        okRic% = (CVDate(valRic) = CVDate(Valore))
                    End If
                Case SS_CELL_TYPE_EDIT, SS_CELL_TYPE_STATIC_TEXT, SS_CELL_TYPE_COMBOBOX
                    okRic% = (RTrim(Valore) Like valRic)
                Case SS_CELL_TYPE_FLOAT, SS_CELL_TYPE_INTEGER
                    If (IsNumeric(Valore)) Then
                        okRic% = CDbl(Valore) = CDbl(valRic)
                    End If
            End Select
            If (okRic%) Then
                DoEvents
                mySS.ReDraw = False
                mySS.OperationMode = SS_OP_MODE_NORMAL
                mySS.Row = Row&
                mySS.Col = ColTrv(ColRic&)
                If mySS.Lock Then
                    mySS.Col = 1
                    Do
                       If Not mySS.Lock Then Exit Do
                       mySS.Col = mySS.Col + 1
                    Loop
                End If
                ssCellActive mySS, mySS.Col, mySS.Row, True     'Anomalia nr. 10073
                mySS.OperationMode = SS_OP_MODE_SINGLE_SELECT
                'mySS.OperationMode = SS_OP_MODE_ROWMODE
                mySS.ReDraw = True
                okRic% = True
                Succ& = Row& + 1
                GoTo fine_Ricerca
            End If
        Next Row&
    End If
    If (Not okRic%) Then
        intMsg = 2124
        Succ& = 1
        On Local Error Resume Next: txtb(0).SetFocus: On Local Error GoTo 0
    End If

fine_Ricerca:
    If intMsg <> 0 Then
        If okRic% Then
            Call MXNU.MsgBoxEX(intMsg, vbExclamation, "")
        Else
            Call MXNU.MsgBoxEX(intMsg, vbExclamation, 1007)
        End If
    End If
    On Local Error GoTo 0
    Go_Ricerca = okRic%
    Exit Function

err_Ricerca:
    Call MXNU.MsgBoxEX(2123, vbExclamation, 1007)
    Resume fine_Ricerca
End Function

Sub Trova_Imposta()

    Dim q As Integer
    Dim Col As Long
    Dim Valore As Variant

    Call CentraFinestra(Me.hwnd)
    cmb.Clear
    cmb.addItem MXNU.CaricaStringaRes(75284)
    For Col = 1 To ColTrv.Count
        mySS.Row = 0
        mySS.Col = ColTrv(Col)
        'valore = mySS.TypeButtonText
        'RIF.S#1404 - tolgo il carattere di sottolineatura &
        Valore = Replace$(ssCellGetValue(mySS, ColTrv(Col), 0), "&", "")
        cmb.addItem Valore
        mySS.Row = 1
        cmb.ItemData(cmb.NewIndex) = mySS.CellType
    Next Col
    If (cmb.ListCount > 0) Then cmb.listIndex = ColOrd
    prevOpMode = mySS.OperationMode
End Sub

Sub txtb_Change(Index As Integer)
    Succ = 1
End Sub

Sub txtb_GotFocus(Index As Integer)
    SelContenuto txtb(0)
End Sub






Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Per Metodo Evolus
If MXNU.ISMETODO2005 Then
    If Not Cancel Then
        On Local Error Resume Next
        If (Not mResize Is Nothing) Then
            mResize.Terminate
            Set mResize = Nothing
        End If
        On Local Error GoTo 0
    End If
End If

End Sub
