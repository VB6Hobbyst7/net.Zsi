VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpreadRigheVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Const ET_RIFERIMENTO = "-1"
Const ET_NEGOZIO = "-2"

Const R_GIACENZA = -2
Const R_OPERATORE = -3
Const R_REPARTO = -4
Const R_TOTNETTORIGAIVATO = -7
Const R_SCONTOFORFAIT = -9

Private WithEvents mSpread As fpSpread
Attribute mSpread.VB_VarHelpID = -1
Private mCGestDoc As MXBusiness.CGestDoc
Private lValidating As Boolean

Event Validazione(ByVal Col As Long, ByVal Row As Long, vValore As Variant, IdCampo As MXBusiness.setCampiRigaDocStd)
Event ValidazionePers(ByVal Col As Long, ByVal Row As Long, vValore As Variant, IdCampo As Long)
Event Selezione(ByVal Col As Long, ByVal Row As Long, IdCampo As MXBusiness.setCampiRigaDocStd)
Event SelezionePers(ByVal Col As Long, ByVal Row As Long, IdCampo As Long)

Event PiedeValidazione(ByVal Col As Long, ByVal Row As Long, vValore As Variant, IdCampo As MXBusiness.setCampiGenPiedeDocStd)
Event PiedeValidazionePers(ByVal Col As Long, ByVal Row As Long, vValore As Variant, IdCampo As Long)
Event PiedeSelezione(ByVal Col As Long, ByVal Row As Long, IdCampo As MXBusiness.setCampiGenPiedeDocStd)
Event PiedeSelezionePers(ByVal Col As Long, ByVal Row As Long, IdCampo As Long)

Private ocDtGen As New cDatiGen
Private xSpread As New CSpread
Private mvarTotImponibile As Double
Private mvarTotIva As Double
Private mvarCodPagCliente As String
Private mvarLstView As ListView
Private mvarfContenitore As Form
Private mvarTipoRiga As String
Private mvarAbilitaSprUm As Boolean
Private mvarPuntoVendita As String
Private WithEvents mvarsPiede As fpSpread
Attribute mvarsPiede.VB_VarHelpID = -1
Private mvarLstViewIVA As ListView

Public Property Set LstViewIVA(ByVal vData As ListView)
    Set mvarLstViewIVA = vData
End Property

Public Property Get LstViewIVA() As ListView
    Set LstViewIVA = mvarLstViewIVA
End Property

Public Property Set sPiede(ByVal vData As fpSpread)
    Set mvarsPiede = vData
End Property

Public Property Get sPiede() As fpSpread
    Set sPiede = mvarsPiede
End Property

Public Property Let PuntoVendita(ByVal vData As String)
    mvarPuntoVendita = vData
End Property

Public Property Get PuntoVendita() As String
    PuntoVendita = mvarPuntoVendita
End Property

Public Property Let AbilitaSprUm(ByVal vData As Boolean)
    mvarAbilitaSprUm = vData
End Property

Public Property Get AbilitaSprUm() As Boolean
    AbilitaSprUm = mvarAbilitaSprUm
End Property

Public Property Let TipoRiga(ByVal vData As String)
    mvarTipoRiga = vData
End Property

Public Property Get TipoRiga() As String
    TipoRiga = mvarTipoRiga
End Property

Public Property Set fContenitore(ByVal vData As Form)
    Set mvarfContenitore = vData
End Property

Public Property Get fContenitore() As Form
    Set fContenitore = mvarfContenitore
End Property

Public Property Set LstView(ByVal vData As ListView)
    Set mvarLstView = vData
End Property

Public Property Get LstView() As ListView
    Set LstView = mvarLstView
End Property

Public Property Let CodPagCliente(ByVal vData As String)
    mvarCodPagCliente = vData
End Property

Public Property Get CodPagCliente() As String
    CodPagCliente = mvarCodPagCliente
End Property

Public Property Let TotIva(ByVal vData As Double)
    mvarTotIva = vData
End Property

Public Property Get TotIva() As Double
    TotIva = mvarTotIva
End Property

Public Property Let TotImponibile(ByVal vData As Double)
    mvarTotImponibile = vData
End Property

Public Property Get TotImponibile() As Double
    TotImponibile = mvarTotImponibile
End Property

Public Function NumRigheDoc() As Long
    On Local Error Resume Next
    NumRigheDoc = mCGestDoc.NumeroRighe
End Function

Public Sub CaricaStruttura()
    Dim i As Long
    Dim cImgBtnUp As String
    Dim cTipoRiga As String
    
    lValidating = True
    
    cTipoRiga = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\" & mvarTipoRiga

    mSpread.ReDraw = False

    mSpread.LoadFromFile cTipoRiga

    mSpread.Protect = True

    cImgBtnUp = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\IMG\SELLSTUP.BMP"

    For i = 1 To mSpread.DataColCnt
        If Not xSpread.ssColHidden(mSpread, i) Then
            mSpread.Col = i
            mSpread.Row = -1

            If mSpread.CellType = CellTypePicture Then

                mSpread.TypePictPicture = LoadPicture(cImgBtnUp)
                'DoEvents

            End If

        End If
    Next i

    LeggiWidthColsDaFile (MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\COL.INI")

    mSpread.ReDraw = True
    
    lValidating = False
End Sub

Public Sub InizializzaEngineDocumento(cConto As String, dData As Date)
    Dim cTipoDoc As String
    Dim i As Long
        
    Call MXNU.ImpostaErroriSuLog(MXNU.GetTempFile, True)
    
    If cConto <> "" Then
    
        If Not mCGestDoc Is Nothing Then
            Call mCGestDoc.Termina
            Set mCGestDoc = Nothing
        End If
        
        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, mvarsPiede, Nothing, Nothing, Nothing)
        
        
        With mCGestDoc
        
            .Stato = GD_INSERIMENTO
            
            If Left(cConto, 1) = "C" Then
                ocDtGen.Chiave = cTipoDocXPrezzi
            Else
                ocDtGen.Chiave = cTipoDocXPrezziFor
            End If
            
            cTipoDoc = ocDtGen.CampoDati
            
            Call .xTDoc.AssegnaCampo("TIPODOC", cTipoDoc)
            Call .xTDoc.AssegnaCampo("ESERCIZIO", MXNU.AnnoAttivo)
            Call .xTDoc.AssegnaCampo("DATADOC", dData)
            Call .xTDoc.AssegnaCampo("CODCLIFOR", cConto)
            
            mvarCodPagCliente = .xTDoc.GrInput("CODPAGAMENTO").ValoreCorrente
        
        End With
        
    End If
    
    Call CalcolaTotali
    
    mvarsPiede.Height = 5985
    mvarsPiede.Width = 4850
    LeggiWidthColsDaFile (MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\COL.INI")
    
    For i = 1 To mSpread.DataColCnt
        If Not xSpread.ssColHidden(mSpread, i) Then
            mSpread.Col = i
            mSpread.Row = -1
            If InStr(",17,18,83,84,126,", "," & Trim(Str(mSpread.GetColItemData(i)) & ",")) > 0 Then
                
                mSpread.CellType = CellTypeFloat
                mSpread.TypeFloatDecimalPlaces = MXNU.DecimaliEuroUnitario
                
            ElseIf InStr(",-7,21,22,85,86,87,88,89,115,116,117,118,119,120,121,122,123,124,", "," & Trim(Str(mSpread.GetColItemData(i)) & ",")) > 0 Then
            
                mSpread.CellType = CellTypeFloat
                mSpread.TypeFloatDecimalPlaces = MXNU.DecimaliEuroTotale
            

            End If

        End If
    Next i

    
    Call MXNU.ChiudiErroriSuLog
    
End Sub

Public Sub TAssegnaCampo(ByVal cCampo As String, vVal As Variant)
    Call mCGestDoc.xTDoc.AssegnaCampo(cCampo, vVal)
End Sub

Public Function TGetCampo(ByVal cCampo As String) As Variant
    TGetCampo = mCGestDoc.xTDoc.GrInput(cCampo).ValoreCorrente
End Function

Public Sub xTraTAssegnaCampo(ByVal cCampo As String, vVal As Variant)
    Call mCGestDoc.xExtraTDoc.AssegnaCampo(cCampo, vVal)
End Sub

Public Function xTraTGetCampo(ByVal cCampo As String) As Variant
    xTraTGetCampo = mCGestDoc.xExtraTDoc.GrInput(cCampo).ValoreCorrente
End Function


Public Property Set Spread(ByVal vData As fpSpread)
    Set mSpread = vData
End Property

Public Property Get Spread() As fpSpread
    Set Spread = mSpread
End Property

Private Sub Class_Terminate()
    
    If Not mCGestDoc Is Nothing Then
        Call mCGestDoc.Termina
        Set mCGestDoc = Nothing
    End If

End Sub

Private Sub mSpread_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nLinkColonna As Long
    mSpread.Col = Col
    mSpread.Row = Row
    If mSpread.CellType = CellTypePicture And Row > 0 Then
    
        If Not lValidating Then
        
            nLinkColonna = mSpread.GetColItemData(Col)
            
            If nLinkColonna > 0 Then
                RaiseEvent Selezione(Col, Row, nLinkColonna)
            Else
                RaiseEvent SelezionePers(Col, Row, nLinkColonna)
            End If
                
        End If
        
        Call CalcolaTotali
    
    End If
    
End Sub

Private Sub mSpread_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim nLinkColonna As Long
    Dim vVar As Variant
    
    If Not lValidating Then
    
        If mSpread.ChangeMade Then
    
            nLinkColonna = mSpread.GetColItemData(Col)
            
            mSpread.GetText Col, Row, vVar
                    
            If nLinkColonna > 0 Then
                RaiseEvent Validazione(Col, Row, vVar, nLinkColonna)
            Else
                RaiseEvent ValidazionePers(Col, Row, vVar, nLinkColonna)
            End If
            
            mSpread.ChangeMade = False
        End If
        
    End If
    
    Call CalcolaTotali
End Sub


Private Sub mvarsPiede_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nLinkColonna As Long
    mvarsPiede.Col = Col
    mvarsPiede.Row = Row
    If Col = 3 Then
    
        If Not lValidating Then
        
            nLinkColonna = mvarsPiede.GetRowItemData(Row)
            
            If nLinkColonna > 0 Then
                RaiseEvent PiedeSelezione(Col, Row, nLinkColonna)
            Else
                RaiseEvent PiedeSelezionePers(Col, Row, nLinkColonna)
            End If
                
        End If
        
        'Call CalcolaTotali
    
    End If
    
End Sub

Private Sub mvarsPiede_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim nLinkColonna As Long
    Dim vVar As Variant
    
    If Not lValidating Then
    
        If mvarsPiede.ChangeMade Then
    
            nLinkColonna = mvarsPiede.GetRowItemData(Row)
            
            mvarsPiede.GetText Col, Row, vVar
                    
            If nLinkColonna > 0 Then
                RaiseEvent PiedeValidazione(Col, Row, vVar, nLinkColonna)
            Else
                RaiseEvent PiedeValidazionePers(Col, Row, vVar, nLinkColonna)
            End If
            
            mvarsPiede.ChangeMade = False
        End If
        
    End If
    
    'Call CalcolaTotali
End Sub


Public Sub PiedeInsertValue(ByVal vVal As Variant, IdCampo As MXBusiness.setCampiGenPiedeDocStd, Optional ByVal lRefresh As Boolean = True)
    Dim i As Long
    Dim nCol As Long
    Dim Row As Long
    Dim vValore As Variant
    Dim lProcedi As Boolean
    Dim vAus As Variant
    
    lValidating = True
    
    On Error Resume Next
    
    Row = TrovaRiga(IdCampo)
    
    StopFormUpdating mvarfContenitore.hwnd
    
    If Not lRefresh Then Call MXNU.ImpostaErroriSuLog(MXNU.GetTempFile, True)
   
    With mCGestDoc
    
        vValore = ""
                
        If vVal <> vValore Then
            
            TAssegnaCampo IdCampo2Campo(IdCampo), vVal
            
        End If
        
        DoEvents
        
        If lRefresh Then
        
            Select Case IdCampo
                Case GP_ANNOTAZIONI:
                    mvarsPiede.SetText 4, Row, MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                Case GP_DATATRASP:
                    ssCellSetValue mvarsPiede, 4, Row, CDate(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                Case GP_ORATRASP:
                    mvarsPiede.SetText 4, Row, MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                Case GP_NUMCOLLI:
                    mvarsPiede.SetInteger 4, Row, MyCInt(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                Case GP_ASPETTOBENI:
                    mvarsPiede.SetText 4, Row, MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                Case GP_BANCAAPP:
                    mvarsPiede.SetText 2, Row, MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                    If objValidBan.Validazione(MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))) Then
                        mvarsPiede.SetText 4, Row, objValidBan.ValoriCampiRit(2)
                    End If
                    
                Case GP_BANCAINC:
                    mvarsPiede.SetText 2, Row, MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                    If objValidBanInc.Validazione(MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))) Then
                        mvarsPiede.SetText 4, Row, objValidBanInc.ValoriCampiRit(2)
                    End If
                    
                Case GP_CODPAG:
                    mvarsPiede.SetText 2, Row, MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                    If objValidPag.Validazione(MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))) Then
                        mvarsPiede.SetText 4, Row, objValidPag.ValoriCampiRit(2)
                    End If

                Case GP_TRASPACURA:
                    mvarsPiede.SetText 2, Row, MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                    If objValidTrasp.Validazione(MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))) Then
                        mvarsPiede.SetText 4, Row, objValidTrasp.ValoriCampiRit(2)
                    End If
                    
                Case GP_CAUSTRASP:
                    mvarsPiede.SetText 2, Row, MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                    If objValidCauTr.Validazione(MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))) Then
                        mvarsPiede.SetText 4, Row, objValidCauTr.ValoriCampiRit(2)
                    End If
                    
                Case GP_PORTO:
                    mvarsPiede.SetText 2, Row, MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))
                    
                    If objValidPorto.Validazione(MyCStr(TGetCampo(IdCampo2Campo(IdCampo)))) Then
                        mvarsPiede.SetText 4, Row, objValidPorto.ValoriCampiRit(2)
                    End If
                    
                Case GP_DATADECORR:
                    ssCellSetValue mvarsPiede, 4, Row, CDate(TGetCampo(IdCampo2Campo(IdCampo)))
                    
            End Select
       
        End If

    End With
    
    'If lRefresh Then Call CalcolaTotali
    If Not lRefresh Then Call MXNU.ChiudiErroriSuLog
    
    lValidating = False
    
    AllowFormUpdating
    
    On Error GoTo 0
    
End Sub

Public Sub InsertValue(ByVal vVal As Variant, IdCampo As MXBusiness.setCampiRigaDocStd, ByVal Col As Long, ByVal Row As Long, Optional ByVal lRefresh As Boolean = True)
    Dim i As Long
    Dim nOldCol As Long
    Dim nOldRow As Long
    Dim vValore As Variant
    Dim lProcedi As Boolean
    Dim vAus As Variant
    
    lValidating = True
    
    On Error Resume Next
    
    nOldCol = mSpread.Col
    nOldRow = mSpread.Row
    
    If lRefresh Then StopFormUpdating mvarfContenitore.hwnd
    
    If Not lRefresh Then Call MXNU.ImpostaErroriSuLog(MXNU.GetTempFile, True)
   
    With mCGestDoc
    
        vValore = "" '.RigaAttiva.ValoreCampo(IdCampo)
        .RigaAttiva.RigaCorr = Row
        DoEvents
        If vVal <> vValore Then
        
            If IdCampo = R_SCONTIESTESI Then
            
                If .RigaAttiva.ValoreCampo(R_CODIVA) >= 100 Then
                    vAus = .RigaAttiva.ValoreCampo(R_PREZZOUNITLORDOIVATO)
                Else
                    vAus = .RigaAttiva.ValoreCampo(R_PREZZOUNITLORDO)
                End If
                .RigaAttiva.ValoreCampo(R_PREZZOUNITLORDO) = 0
                .RigaAttiva.ValoreCampo(R_PREZZOUNITLORDO) = vAus
            End If
            
            .RigaAttiva.ValoreCampo(IdCampo) = vVal
            
        End If
        
'        DoEvents
        
        If lRefresh Then
            
            For i = 1 To mSpread.DataColCnt
                mSpread.Col = i
                
                If mSpread.CellType <> CellTypePicture Then
                
                    Select Case mSpread.CellType
                        Case CellTypeDate:
                            mSpread.SetText i, Row, CStr(.RigaAttiva.ValoreCampo(mSpread.GetColItemData(i)))
                        Case CellTypeEdit:
                            mSpread.SetText i, Row, MyCStr(.RigaAttiva.ValoreCampo(mSpread.GetColItemData(i)))
                        Case CellTypeFloat:
                            mSpread.SetFloat i, Row, MyCDbl(.RigaAttiva.ValoreCampo(mSpread.GetColItemData(i)))
                        Case CellTypeInteger:
                            mSpread.SetInteger i, Row, MyCDbl(.RigaAttiva.ValoreCampo(mSpread.GetColItemData(i)))
                        Case CellTypeCheckBox:
                            mSpread.SetInteger i, Row, MyCInt(.RigaAttiva.ValoreCampo(mSpread.GetColItemData(i)))
                    End Select
                    
                    DoEvents
                    
                End If
            
            Next i
        
        End If

    End With
    
    If lRefresh Then Call CalcolaTotali
    If Not lRefresh Then Call MXNU.ChiudiErroriSuLog
    
'    mSpread.Row = nOldRow
'    mSpread.Col = nOldCol
    
    lValidating = False
    
    If lRefresh Then AllowFormUpdating
    
    On Error GoTo 0
    
End Sub

Public Function SpreadCompilato() As Boolean
    Dim i As Long
    Dim nArt As Long
    Dim nPrz As Long
    Dim nQta As Long
    
    Dim nCArt As Long
    Dim nCPrz As Long
    Dim nCQta As Long

    For i = 1 To mSpread.DataColCnt
    
        If mSpread.GetColItemData(i) = R_PREZZOUNITLORDOIVATO Then
            nCPrz = i
        End If
        
        If mSpread.GetColItemData(i) = R_CODART Then
            nCArt = i
        End If
        
        If mSpread.GetColItemData(i) = R_QTAGEST Then
            nCQta = i
        End If
        
    Next i
    
    SpreadCompilato = False
    
    nArt = 0
    nPrz = 0
    nQta = 0
    
    For i = 1 To mSpread.DataRowCnt
        If xSpread.ssCellGetValue(mSpread, nCArt, i) <> "" Then
            nArt = nArt + 1
        End If
        
        If xSpread.ssCellGetValue(mSpread, nCPrz, i) <> 0 Then
            nPrz = nPrz + 1
        End If
        
        If xSpread.ssCellGetValue(mSpread, nCQta, i) <> 0 Then
            nQta = nQta + 1
        End If
        
    Next
    
    If nArt + nPrz + nQta > 0 Then
        SpreadCompilato = True
    End If

End Function

Public Function TrovaColonna(IdCampo As MXBusiness.setCampiRigaDocStd) As Long
    Dim i As Long
    
    For i = 1 To mSpread.DataColCnt
    
        If mSpread.GetColItemData(i) = IdCampo Then
            TrovaColonna = i
        End If
        
    Next i
    
End Function

Public Function TrovaRiga(IdCampo As MXBusiness.setCampiGenPiedeDocStd) As Long
    Dim i As Long
    
    For i = 1 To mvarsPiede.DataRowCnt
    
        If mvarsPiede.GetRowItemData(i) = IdCampo Then
            TrovaRiga = i
        End If
        
    Next i
    
End Function


Public Function GetValue(IdCampo As MXBusiness.setCampiRigaDocStd, ByVal Row As Long) As Variant
    Dim i As Long
    Dim nOldCol As Long
    Dim nOldRow As Long
    Dim vValore As Variant
    Dim lProcedi As Boolean
    Dim nCol As Long
    
    With mCGestDoc
    
        'nOldRow = .RigaAttiva.RigaCorr
        
        .RigaAttiva.RigaCorr = Row
        
        GetValue = .RigaAttiva.ValoreCampo(IdCampo)
        
       ' .RigaAttiva.RigaCorr = nOldRow

    End With
    
End Function

Public Sub CalcolaTotali()
    With mCGestDoc
        
        'StopFormUpdating mvarfContenitore.hwnd
        
        Call .Calcolo_Totali
        
        mvarTotImponibile = .xTDoc.GrInput("TOTIMPONIBILEEURO").ValoreCorrente
        mvarTotIva = .xTDoc.GrInput("TOTIMPOSTAEURO").ValoreCorrente
        
        Call FillSprIva
        
        If mvarAbilitaSprUm Then Call FillSprQta
        
        
        
        'AllowFormUpdating
        
    End With
End Sub

Public Function CancellaRiga(ByVal nRiga As Long)
    Dim cMsg As String
    
    If ssCellGetValue(mSpread, 0, nRiga) = "A" Then
    
        mCGestDoc.RigaAttiva.RigaCorr = nRiga
        mCGestDoc.RigaAttiva.AnnullaRiga True
        Call ssCellSetValue(mSpread, 0, nRiga, nRiga)
        Call CalcolaTotali
    Else
        cMsg = "Confermi l'eliminazione della riga :r ?"
        cMsg = Replace(cMsg, ":r", nRiga)
        
        If MXNU.MsgBoxEX(cMsg, vbYesNo, "Conferma Eliminazione") = vbYes Then
            mCGestDoc.RigaAttiva.RigaCorr = nRiga
            mCGestDoc.RigaAttiva.AnnullaRiga
            Call ssCellSetValue(mSpread, 0, nRiga, "A")
            Call CalcolaTotali
        End If
    End If
End Function

Public Function CalcolaNettoIVA(ByVal nRiga) As Double
    Dim nAlIva As Double
    Dim cSql As String
    Dim rSql As CRecordSet
    
    CalcolaNettoIVA = GetValue(R_PREZZOUNITLORDOIVATO, nRiga) * ((100 - GetValue(R_SCONTORIGA, nRiga)) / 100) * GetValue(R_QTAGEST, nRiga)
    
End Function

Public Sub SalvaWidthColsSuFile(NomeFile As String)
   Dim i As Long
   Dim clsIni As cIniFiles
   Dim intFile As Integer
   Dim lngCurCol As Long
   '
   On Error GoTo errhand
   '
   If Exist(NomeFile) Then Kill (NomeFile) ' questo perchè potrei decidere di togliere/mettere colonne
   Set clsIni = New cIniFiles
   With mSpread
      lngCurCol = .Col
      .UnitType = UnitTypeTwips
      For i = 1 To .MaxCols
         .Col = i
         clsIni.SetValue NomeFile, "ColWidths", "Col" + Trim$(Str$(i)), .ColWidth(i)
      Next
      .Col = lngCurCol
   End With
   
   SaveSetting App.Title, "SPRQTA", "NRIGHE", mvarLstView.ColumnHeaders(1).Width
   SaveSetting App.Title, "SPRQTA", "UM", mvarLstView.ColumnHeaders(2).Width
   SaveSetting App.Title, "SPRQTA", "QTA", mvarLstView.ColumnHeaders(3).Width
   
   SaveSetting App.Title, "SPRIVA", "IVA", mvarLstViewIVA.ColumnHeaders(1).Width
   SaveSetting App.Title, "SPRIVA", "DSC", mvarLstViewIVA.ColumnHeaders(2).Width
   SaveSetting App.Title, "SPRIVA", "IMP", mvarLstViewIVA.ColumnHeaders(3).Width
   SaveSetting App.Title, "SPRIVA", "STA", mvarLstViewIVA.ColumnHeaders(4).Width
   

cleanexit:
   If Not clsIni Is Nothing Then Set clsIni = Nothing
   Exit Sub

errhand:
   
End Sub

Public Sub LeggiWidthColsDaFile(NomeFile As String)
   Dim i As Long
   Dim clsIni As cIniFiles
   Dim lngCurCol As Long
   Dim lngWidth As Long
   '
   On Error GoTo errhand
   Set clsIni = New cIniFiles
   With mSpread
      lngCurCol = .Col
      .UnitType = UnitTypeTwips
      For i = 1 To .MaxCols
         .Col = i
         lngWidth = clsIni.GetValue(NomeFile, "ColWidths", "Col" + Trim$(Str$(i)), IIf(i = 0, 500, 1000))
         .ColWidth(i) = lngWidth
      Next
      .Col = lngCurCol
   End With
   
   mvarLstView.ColumnHeaders(1).Width = GetSetting(App.Title, "SPRQTA", "NRIGHE", mvarLstView.ColumnHeaders(1).Width)
   mvarLstView.ColumnHeaders(2).Width = GetSetting(App.Title, "SPRQTA", "UM", mvarLstView.ColumnHeaders(2).Width)
   mvarLstView.ColumnHeaders(3).Width = GetSetting(App.Title, "SPRQTA", "QTA", mvarLstView.ColumnHeaders(3).Width)

   mvarLstViewIVA.ColumnHeaders(1).Width = GetSetting(App.Title, "SPRIVA", "IVA", mvarLstViewIVA.ColumnHeaders(1).Width)
   mvarLstViewIVA.ColumnHeaders(2).Width = GetSetting(App.Title, "SPRIVA", "DSC", mvarLstViewIVA.ColumnHeaders(2).Width)
   mvarLstViewIVA.ColumnHeaders(3).Width = GetSetting(App.Title, "SPRIVA", "IMP", mvarLstViewIVA.ColumnHeaders(3).Width)
   mvarLstViewIVA.ColumnHeaders(4).Width = GetSetting(App.Title, "SPRIVA", "STA", mvarLstViewIVA.ColumnHeaders(4).Width)
   

cleanexit:
   If Not clsIni Is Nothing Then Set clsIni = Nothing
   Exit Sub
   
errhand:
   Resume cleanexit
End Sub

Public Function CreaDoc(ByVal cTipoOperazione As String, ssRighe As fpSpread) As Long
    Dim nRiga As Long
    Dim nCol As Long
    Dim vVar As Variant
    Dim intNewEsercizio As Integer
    Dim lngNewNrDoc As Long
    Dim strNewBis As String
    Dim nNumRiga As Long
    Dim n As Long
    Dim i As Long
    Dim cRif As String
    Dim nCauMag As String
    
    Call MXNU.ImpostaErroriSuLog(MXNU.GetTempFile, True)
    
    With mCGestDoc
    
        nCauMag = GetCausaleMag(ListaOperazioni(cTipoOperazione).TipoDocumento)
    
        cRif = .xExtraTDoc.GrInput(ET_RIFERIMENTO).ValoreCorrente
    
        Call .xTDoc.AssegnaCampo("TIPODOC", ListaOperazioni(cTipoOperazione).TipoDocumento)
        
        Call .xExtraTDoc.AssegnaCampo(ET_RIFERIMENTO, cRif)
        Call .xExtraTDoc.AssegnaCampo(ET_NEGOZIO, mvarPuntoVendita)
        
        For nRiga = 1 To mvarsPiede.DataRowCnt
        
            Select Case mvarsPiede.GetRowItemData(nRiga)
                Case 1, 2, 6, 13, 15, 16:
                    TAssegnaCampo IdCampo2Campo(mvarsPiede.GetRowItemData(nRiga)), ssCellGetValue(mvarsPiede, 2, nRiga)
                Case Else
                    TAssegnaCampo IdCampo2Campo(mvarsPiede.GetRowItemData(nRiga)), ssCellGetValue(mvarsPiede, 4, nRiga)
            End Select
            
        Next nRiga
    
        For nRiga = 1 To .NumeroRighe

            .RigaAttiva.RigaCorr = nRiga
            
            .RigaAttiva.ValoreCampo(R_CAUSMAG) = nCauMag

            If ssCellGetValue(ssRighe, 0, nRiga) <> "A" Then

                If (ListaOperazioni(cTipoOperazione).DocTipo = "B" Or ListaOperazioni(cTipoOperazione).DocTipo = "F") And ListaOperazioni(cTipoOperazione).TipoOperazione <> "NC" Then

                    If .RigaAttiva.ValoreCampo(R_CODIVA) >= 100 Then

                        .RigaAttiva.ValoreCampo(R_CODIVA) = .RigaAttiva.ValoreCampo(R_CODIVA) - 100

                    End If

                End If

            Else

                .RigaAttiva.AnnullaRiga

            End If

        Next nRiga
            
        Call .Calcolo_Totali
        
        Call MXNU.ChiudiErroriSuLog
        
        intNewEsercizio = .xTDoc.GrInput("ESERCIZIO").ValoreCorrente
        lngNewNrDoc = .xTDoc.GrInput("NUMERODOC").ValoreCorrente
        strNewBis = .xTDoc.GrInput("BIS").ValoreCorrente
        
        CreaDoc = 0
        
        If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis, MXNU.LeggiProfilo(MXNU.File_ini_Documenti, "CREAMOVMAG", "BATCH", "1"), MXNU.LeggiProfilo(MXNU.File_ini_Documenti, "CREATRANSITORIO", "BATCH", "1")) Then
            CreaDoc = .xTDoc.GrInput("PROGRESSIVO").ValoreCorrente
        Else
            Call MXNU.MsgBoxEX("Errore nella generazione del documento: " & ListaOperazioni(cTipoOperazione).TipoDocumento & "/" & intNewEsercizio & "/" & lngNewNrDoc, vbCritical, "Errore durante la generazione del documento!")
        End If
        
    End With
    
End Function

Public Sub CaricaDoc(ByVal nIdTesta As Long, cTpOp As String, ssRighe As fpSpread)
    Dim vbDoc As MXBusiness.CGestDoc
    Dim nRiga As Long
    Dim nCol As Long
    Dim vVar As Variant
    Dim intNewEsercizio As Integer
    Dim lngNewNrDoc As Long
    Dim strNewBis As String
    Dim nNumRiga As Long
    Dim nOldLordo As Double
    
    If Not mCGestDoc Is Nothing Then
        Call mCGestDoc.Termina
        Set mCGestDoc = Nothing
    End If
    
    Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
    
    With mCGestDoc
        Call .xTDoc.AssegnaCampo("PROGRESSIVO", nIdTesta)
        
        intNewEsercizio = .xTDoc.GrInput("ESERCIZIO").ValoreCorrente
        lngNewNrDoc = .xTDoc.GrInput("NUMERODOC").ValoreCorrente
        strNewBis = .xTDoc.GrInput("BIS").ValoreCorrente
    
        .Stato = GD_MODIFICA
        
        nNumRiga = 0
        
        ssRighe.ReDraw = False
        
        For nRiga = 1 To .NumeroRighe

            .RigaAttiva.RigaCorr = nRiga

            ssRighe.Row = nRiga

            For nCol = 1 To ssRighe.DataColCnt

                ssRighe.Col = nCol

                If ssRighe.CellType <> CellTypePicture Then
                
                    If ssRighe.GetColItemData(nCol) <> R_TIPORIGA And ssRighe.GetColItemData(nCol) <> R_RIGAPRELEVABILE Then
                    
                        If ssRighe.GetColItemData(nCol) = R_PREZZOUNITLORDO Then
                        
                            If ListaOperazioni(cTpOp).TipoOperazione = "NC" Then

                                If .RigaAttiva.ValoreCampo(R_CODIVA) < 100 Then
                                    
                                    nOldLordo = .RigaAttiva.ValoreCampo(R_PREZZOUNITLORDOIVATO)
                                    
                                    .RigaAttiva.ValoreCampo(R_CODIVA) = .RigaAttiva.ValoreCampo(R_CODIVA) + 100
                                    
                                    .RigaAttiva.ValoreCampo(R_PREZZOUNITLORDO) = nOldLordo

                                    'oRigheVB.InsertValue .RigaAttiva.ValoreCampo(R_CODIVA), R_CODIVA, nCol, nRiga, False

                                End If

                            End If
                           
                        End If
                        
                        Select Case ssRighe.CellType
                            Case CellTypeDate:
                                ssRighe.SetText nCol, nRiga, CStr(.RigaAttiva.ValoreCampo(ssRighe.GetColItemData(nCol)))
                            Case CellTypeEdit:
                                ssRighe.SetText nCol, nRiga, MyCStr(.RigaAttiva.ValoreCampo(ssRighe.GetColItemData(nCol)))
                            Case CellTypeFloat:
                                ssRighe.SetFloat nCol, nRiga, MyCDbl(.RigaAttiva.ValoreCampo(ssRighe.GetColItemData(nCol)))
                            Case CellTypeInteger:
                                ssRighe.SetInteger nCol, nRiga, MyCDbl(.RigaAttiva.ValoreCampo(ssRighe.GetColItemData(nCol)))
                            Case CellTypeCheckBox:
                                ssRighe.SetInteger nCol, nRiga, MyCInt(.RigaAttiva.ValoreCampo(ssRighe.GetColItemData(nCol)))
                        End Select
                        
                    End If

                End If

                'DoEvents

            Next nCol
            
            DoEvents

        Next nRiga
        ssRighe.ReDraw = True
        
        For nRiga = 1 To mvarsPiede.DataRowCnt

            If Not xSpread.ssRowHidden(mvarsPiede, nRiga) Then

                PiedeInsertValue TGetCampo(IdCampo2Campo(mvarsPiede.GetRowItemData(nRiga))), mvarsPiede.GetRowItemData(nRiga)
             
                DoEvents

            End If

        Next nRiga
        
        Call MXNU.ChiudiErroriSuLog
        
    End With
    
    Call CalcolaTotali

End Sub

Private Sub FillSprQta()
    Dim i As Long
    Dim li As ListItem
    Dim oSprQta As New cCollSprQta
    Dim oItem As cSprQta
    
    For i = 1 To mCGestDoc.NumeroRighe

        If ssCellGetValue(mSpread, 0, i) <> "A" And GetValue(R_UMGEST, i) <> "" Then

            If oSprQta(GetValue(R_UMGEST, i)) Is Nothing Then

                oSprQta.Add 1, GetValue(R_UMGEST, i), MyCDbl(GetValue(R_QTAGEST, i)), MyCDbl(GetValue(R_TOTNETTORIGA, i)), MyCDbl(GetValue(-7, i)), GetValue(R_UMGEST, i)

            Else

                oSprQta(GetValue(R_UMGEST, i)).NRighe = oSprQta(GetValue(R_UMGEST, i)).NRighe + 1
                oSprQta(GetValue(R_UMGEST, i)).Qta = oSprQta(GetValue(R_UMGEST, i)).Qta + MyCDbl(GetValue(R_QTAGEST, i))
                'oSprQta(GetValue(R_UMGEST, i)).TotNetto = oSprQta(GetValue(R_UMGEST, i)).TotNetto + MyCDbl(GetValue(R_TOTNETTORIGA, i))
                'oSprQta(GetValue(R_UMGEST, i)).TotIvato = oSprQta(GetValue(R_UMGEST, i)).TotIvato + MyCDbl(GetValue(-7, i))

            End If

        End If

    Next i
    
    mvarLstView.ListItems.Clear
    
    For Each oItem In oSprQta
    
        Set li = mvarLstView.ListItems.Add(, , oItem.UM, Empty, Empty)
        li.SubItems(1) = oItem.NRighe
        li.SubItems(2) = Format(oItem.Qta, MkPictDecimal(MXNU.DecimaliEuroTotale))
        'li.SubItems(3) = oItem.TotNetto
        'li.SubItems(4) = oItem.TotIvato
    
    Next
    
    Set oSprQta = Nothing
End Sub

Private Sub FillSprIva()
    Dim i As Long
    Dim li As ListItem
    
    mvarLstViewIVA.ListItems.Clear
    
    For i = 1 To mCGestDoc.Castelletto.Count
        
        If mCGestDoc.Castelletto(i).CodIVA <> "000" Then
        
            Set li = mvarLstViewIVA.ListItems.Add(, , mCGestDoc.Castelletto(i).Aliq, Empty, Empty)
            li.SubItems(1) = mCGestDoc.Castelletto(i).DscIVA
            li.SubItems(2) = Format(mCGestDoc.Castelletto(i).Imponib(CTD_EURO, VI_INIZIALE), MkPictDecimal(MXNU.DecimaliEuroTotale))
            li.SubItems(3) = Format(mCGestDoc.Castelletto(i).Imposta(CTD_EURO, VI_INIZIALE), MkPictDecimal(MXNU.DecimaliEuroTotale))
            li.SubItems(4) = Format(mCGestDoc.Castelletto(i).Imponib(CTD_EURO, VI_INIZIALE) + mCGestDoc.Castelletto(i).Imposta(CTD_EURO, VI_INIZIALE), MkPictDecimal(MXNU.DecimaliEuroTotale))
        
        End If
    
    Next i
    
End Sub

Public Function IdCampo2Campo(ByVal IdCampo As MXBusiness.setCampiGenPiedeDocStd) As String

    Select Case IdCampo
        Case GP_ANNOTAZIONI:
            IdCampo2Campo = "ANNOTAZIONI"
        Case GP_DATATRASP:
            IdCampo2Campo = "DATAINIZIOTRASP"
        Case GP_ORATRASP:
            IdCampo2Campo = "ORAINIZIOTRASP"
        Case GP_NUMCOLLI:
            IdCampo2Campo = "NUMCOLLI"
        Case GP_ASPETTOBENI:
            IdCampo2Campo = "ASPETTOBENI"
        Case GP_BANCAAPP:
            IdCampo2Campo = "CODBANCAAPPOGGIO"
        Case GP_BANCAINC:
            IdCampo2Campo = "CODBANCAINCASSO"
        Case GP_TRASPACURA:
            IdCampo2Campo = "TRASPACURA"
        Case GP_CAUSTRASP:
            IdCampo2Campo = "CAUSALETRASP"
        Case GP_PORTO:
            IdCampo2Campo = "PORTO"
        Case GP_DATADECORR:
            IdCampo2Campo = "DATADEC"
        Case GP_CODPAG:
            IdCampo2Campo = "CODPAGAMENTO"
    End Select

End Function

Private Function GetCausaleMag(ByVal cTipoDoc As String) As Integer
    Dim cSql As String
    Dim rSql As CRecordSet

    cSql = "SELECT CAUSALEMAG FROM PARAMETRIDOC WHERE CODICE = ':CD'"
    cSql = Replace(cSql, ":CD", cTipoDoc)

    Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)

    GetCausaleMag = MXDB.dbGetCampo(rSql, TIPO_SNAPSHOT, "CAUSALEMAG", 0)

    Call MXDB.dbChiudiSS(rSql)
    
End Function

