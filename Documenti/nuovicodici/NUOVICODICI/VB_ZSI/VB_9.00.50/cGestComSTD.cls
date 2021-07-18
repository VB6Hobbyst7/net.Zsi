VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGestComSTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe che stampa i documenti preventivo della SIVIM da Word"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarmBis As Variant 'Copia locale.
Private mvarmNumeroDoc As Variant 'Copia locale.
Private mvarmCommessa As Variant 'Copia locale.
Private mvarmTipoDoc As Variant 'Copia locale.
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarmProgCommessa As Variant 'Copia locale.
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarmNomeFile As Variant 'Copia locale.

Function FrmtValore(ByVal valore As Variant, ByVal TipoValore As Integer) As String
    ' funzione che consente di impostare "" nel caso il valore passato sia 0
    ' oppure il valore con il formato corretto Quantità=0/Importo Unitario=1/Importo Totale=2
    If valore = "0" Then
        FrmtValore = ""
    Else
        Select Case TipoValore
            Case 0 ' quantità
                FrmtValore = mxutil.FormatCampo(valore, MXNU.FORMATO_QUANTITA)
            Case 1 ' unitario
                FrmtValore = mxutil.FormatCampo(valore, MXNU.FORMATO_EURO_UNITARIO)
            Case 2 ' totale
                FrmtValore = mxutil.FormatCampo(valore, MXNU.FORMATO_EURO_TOTALE)
            Case Else
                FrmtValore = valore
        End Select
    End If
End Function

Public Property Let mNomeFile(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.mNomeFile = 5
    mvarmNomeFile = vData
End Property


Public Property Set mNomeFile(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un oggetto a una proprietà, a sinistra di un'istruzione Set.
'Syntax: Set x.mNomeFile = Form1
    Set mvarmNomeFile = vData
End Property


Public Property Get mNomeFile() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.mNomeFile
    If IsObject(mvarmNomeFile) Then
        Set mNomeFile = mvarmNomeFile
    Else
        mNomeFile = mvarmNomeFile
    End If
End Property



Public Property Let mProgCommessa(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.mProgCommessa = 5
    mvarmProgCommessa = vData
End Property


Public Property Set mProgCommessa(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un oggetto a una proprietà, a sinistra di un'istruzione Set.
'Syntax: Set x.mProgCommessa = Form1
    Set mvarmProgCommessa = vData
End Property


Public Property Get mProgCommessa() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.mProgCommessa
    If IsObject(mvarmProgCommessa) Then
        Set mProgCommessa = mvarmProgCommessa
    Else
        mProgCommessa = mvarmProgCommessa
    End If
End Property



Public Property Let mTipoDoc(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.mTipoDoc = 5
    mvarmTipoDoc = vData
End Property


Public Property Set mTipoDoc(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un oggetto a una proprietà, a sinistra di un'istruzione Set.
'Syntax: Set x.mTipoDoc = Form1
    Set mvarmTipoDoc = vData
End Property


Public Property Get mTipoDoc() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.mTipoDoc
    If IsObject(mvarmTipoDoc) Then
        Set mTipoDoc = mvarmTipoDoc
    Else
        mTipoDoc = mvarmTipoDoc
    End If
End Property



Public Property Let mCommessa(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.mCommessa = 5
    mvarmCommessa = vData
End Property


Public Property Set mCommessa(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un oggetto a una proprietà, a sinistra di un'istruzione Set.
'Syntax: Set x.mCommessa = Form1
    Set mvarmCommessa = vData
End Property


Public Property Get mCommessa() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.mCommessa
    If IsObject(mvarmCommessa) Then
        Set mCommessa = mvarmCommessa
    Else
        mCommessa = mvarmCommessa
    End If
End Property



Public Property Let mNumeroDoc(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.mNumeroDoc = 5
    mvarmNumeroDoc = vData
End Property


Public Property Set mNumeroDoc(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un oggetto a una proprietà, a sinistra di un'istruzione Set.
'Syntax: Set x.mNumeroDoc = Form1
    Set mvarmNumeroDoc = vData
End Property


Public Property Get mNumeroDoc() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.mNumeroDoc
    If IsObject(mvarmNumeroDoc) Then
        Set mNumeroDoc = mvarmNumeroDoc
    Else
        mNumeroDoc = mvarmNumeroDoc
    End If
End Property



Public Property Let mBis(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.mBis = 5
    mvarmBis = vData
End Property


Public Property Set mBis(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un oggetto a una proprietà, a sinistra di un'istruzione Set.
'Syntax: Set x.mBis = Form1
    Set mvarmBis = vData
End Property


Public Property Get mBis() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.mBis
    If IsObject(mvarmBis) Then
        Set mBis = mvarmBis
    Else
        mBis = mvarmBis
    End If
End Property



Public Function StampaPrev(ByVal strNomeDOT As String, ByVal intTipoPrev As Integer, strTipoPrev As String, ByVal lngNrPrev As Long, ByVal strBis As String, ByVal strEsercizio As String _
                , ByVal bolAnt As Boolean, intTipo As Integer, lngRifPr As Long, strTipoP As String, bolRicavi As Boolean, ByVal strPathWORD As String)
    
    Dim NRighe As Variant
    Dim strSql As String
    Dim rSql As CRecordSet
    Dim hSSSrc As CRecordSet
    Dim rRigheCT As CRecordSet  ' recordset con le righe per il preventivo CONTO TERZI
    Dim bolEnd As Boolean
    ' testata documento
    Dim Cliente As String
    Dim DataDoc As Date
    Dim nrDoc As String
    Dim TipoDoc As String
    Dim strDoc As String
    Dim strIndirizzo As String
    Dim strLocalita As String
    Dim strDscCliente As String
    Dim strDSCTipoPrev As String
    Dim strDSCPag As String
    ' righe documento
    Dim CodArt As String
    Dim CodStampa As String
    Dim CodGruppo As String
    Dim CodSottoGruppo As String
    Dim DscArt As String
    Dim DscEstesa As String
    Dim Sottocommessa As String
    Dim intExtra As Integer
    Dim Lunghezza As String
    Dim Larghezza As String
    Dim Altezza As String
    Dim Numero As String
    Dim dblQta As String
    Dim UM As String
    Dim dblImp1 As String
    Dim dblImp2 As String
    Dim strNomeFile As String
    Dim strUnitaImm As String       ' unita immobiliare
    
    Dim CostoUnit As String      ' costo unitario senza maggiorazione
    Dim PercMaggC As String      ' percentuale di maggiorazione costi
    Dim CostoUnitTot As String   ' costo unitario con maggiorazione
    Dim CostoTot As String       ' costo totale
    
    Dim PercMaggR As String      ' percentuale di maggiorazione costi
    Dim RicavoUnitTot As String   ' Ricavo unitario con maggiorazione
    Dim RicavoTotLordo As String      ' Ricavo totale con maggiorazione
    Dim PercSconto As String      ' percentuale di sconto
    Dim RicavoTotNetto As String      ' Ricavo totale con sconto
    Dim PercSicurezza As String      ' percentuale di maggiorazione sicurezza
    Dim RicavoSicurezza As String       ' ricavo sicurezza
    Dim RicavoTot As String      ' ricavo totale riga
    
    Dim AppWord As Object 'New Word.Application
    Dim DocWord As Object
    Dim strModelloDot As String
    Dim dblTotImp As Double
    Dim dblTotRic As Double
    Dim I As Integer
    Dim lngCP As Long       ' posizione della riga
    
' ObjInserter
'---------------------------------------------
    On Error GoTo ComProcedi_Click_ErrHandler
'---------------------------------------------
    
    Set AppWord = CreateObject("Word.application")
            
    strSql = "SELECT * FROM VISTA_STPDOCPREV WHERE "
    strSql = strSql & " TIPODOC  = :TIPODOC AND NUMERODOC = :NUMERODOC AND BIS = :BIS AND ESERCIZIO = :ESERCIZIO"
    strSql = Replace(strSql, ":TIPODOC", hndDBArchivi.FormatoSQL(strTipoPrev, DB_TEXT))
    strSql = Replace(strSql, ":NUMERODOC", hndDBArchivi.FormatoSQL(lngNrPrev, DB_LONG))
    strSql = Replace(strSql, ":BIS", hndDBArchivi.FormatoSQL(strBis, DB_TEXT))
    strSql = Replace(strSql, ":ESERCIZIO", hndDBArchivi.FormatoSQL(strEsercizio, DB_INTEGER))
    Set hSSSrc = MXDB.dbCreaSS(hndDBArchivi, strSql)
    


    With MXDB
        If .dbFineTab(hSSSrc) Then

            Exit Function
        Else
            strDscCliente = .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "DSCCONTO1", "")
            strIndirizzo = .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "INDIRIZZO", "")
            strLocalita = .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "CAP", "") & " -" & .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "LOCALITA", "") & " - " & .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "PROVINCIA", "")
            strDSCTipoPrev = .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "dscTipoPrev", "")
            nrDoc = .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "numerodoc", "")
            If strBis <> "" And strBis <> " " Then
                nrDoc = nrDoc & "/" & strBis
            End If
            strDSCPag = .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "DSCPAG", "")
            If .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "DSCPAGEXTRA", "") <> "" Then
                strDSCPag = .dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "DSCPAGEXTRA", "")
            End If
        End If
    End With


        ' Setta proprietà Word
        AppWord.Visible = bolAnt

        
        ' Apre file word x stampa documento
        ' Apre documento word
        Set DocWord = AppWord.Documents.Open(strNomeDOT)
        ' Ritardo
        For I = 1 To 10000
        Next I
        
        ' dati TESTATA
        ' Assegna dato al campo 1 del modulo word
        
        If DocWord.Bookmarks.Exists("TipoDoc") Then
            'If bolRicavi And intTipoPrev <> 1 Then
            Dim strDir As String
            Dim strFil As String
            Call mxutil.scomponiPathFile(strNomeDOT, strDir, strFil)
            If Left(strFil, 6) = "PCCPRC" Then
                DocWord.Bookmarks("TipoDoc").Range.Text = "PREVENTIVO COSTI E RICAVI" 'strDSCTipoPrev
            Else
                DocWord.Bookmarks("TipoDoc").Range.Text = strDSCTipoPrev
            End If
        End If
        
        ' Assegna dato al campo 1 del modulo word
        If DocWord.Bookmarks.Exists("DataPrev") Then
            DocWord.Bookmarks("DataPrev").Range.Text = Date
        End If
        
        ' Assegna dato al campo 1 del modulo word
        If DocWord.Bookmarks.Exists("Pagamento") Then
            DocWord.Bookmarks("Pagamento").Range.Text = strDSCPag
        End If
        
        ' Assegna dato al campo 1 del modulo word
        If DocWord.Bookmarks.Exists("NrPrev") Then
            DocWord.Bookmarks("NrPrev").Range.Text = nrDoc
        End If
        
        ' Assegna dato al campo 2 del modulo word
        If DocWord.Bookmarks.Exists("Cliente") Then
            DocWord.Bookmarks("Cliente").Range.Text = strDscCliente
        End If
    
        ' Assegna dato al campo 2 del modulo word
        If DocWord.Bookmarks.Exists("Indirizzo") Then
            DocWord.Bookmarks("Indirizzo").Range.Text = strIndirizzo
        End If
        
        ' Assegna dato al campo 2 del modulo word
        If DocWord.Bookmarks.Exists("Localita") Then
            DocWord.Bookmarks("Localita").Range.Text = strLocalita
        End If
        
       
        If intTipo = 1 And intTipoPrev = 1 Then
            ' conto terzi
            strSql = "SELECT * FROM VISTA_GENERAPr1 WHERE IDSESSIONE = :IDSESSIONE and RIFPROGRESSIVO = :RIFPROGRESSIVO ORDER BY NRRIGA"
            strSql = Replace(strSql, ":IDSESSIONE", hndDBArchivi.FormatoSQL(MXNU.IDSessione, DB_LONG))
            strSql = Replace(strSql, ":RIFPROGRESSIVO", hndDBArchivi.FormatoSQL(lngRifPr, DB_LONG))
            
            With MXDB
                Set rRigheCT = .dbCreaSS(hndDBArchivi, strSql)
                
                Do While Not .dbFineTab(rRigheCT)
                    lngCP = rRigheCT.RecSet.AbsolutePosition + 2
                    CodStampa = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODSTAMPA", ""))
                    CodArt = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODART", ""))
                    'CodGruppo = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODGRUPPO", ""))
                    'CodSottoGruppo = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODSOTTOGRUPPO", ""))
                    DscArt = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "DESCRIZIONE", ""))
                    'DscEstesa = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "DESCRESTESA", ""))
                    dblQta = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "QTA", 0))
                    UM = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "UM", ""))
                    RicavoTot = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOTOT", 0))
                    
                    If CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODSTAMPA", "")) <> "" Then
                        DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodStampa
                    Else
                        DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodArt
                    End If
                    DocWord.tables(1).Cell(lngCP, 2).Range.Text = DscArt
                    DocWord.tables(1).Cell(lngCP, 3).Range.Text = FrmtValore(RicavoTot, 2)
                    dblTotImp = dblTotImp + CDbl(RicavoTot)
                    Call .dbSuccessivo(rRigheCT)
                    If Not .dbFineTab(rRigheCT) Then
                        Call DocWord.tables(1).Rows.Add
                    End If
                Loop
                
                Call .dbChiudiSS(rRigheCT)
                Set rRigheCT = Nothing
            End With
        
        ElseIf intTipo = 2 And intTipoPrev = 1 Then
            ' a corpo
            strSql = "SELECT * FROM VISTA_GENERAPCC WHERE IDSESSIONE = :IDSESSIONE and RIFPROGRESSIVO = :RIFPROGRESSIVO"
            strSql = Replace(strSql, ":IDSESSIONE", hndDBArchivi.FormatoSQL(MXNU.IDSessione, DB_LONG))
            strSql = Replace(strSql, ":RIFPROGRESSIVO", hndDBArchivi.FormatoSQL(lngRifPr, DB_LONG))
            
            With MXDB
                Set rRigheCT = .dbCreaSS(hndDBArchivi, strSql)
                
                Do While Not .dbFineTab(rRigheCT)
                    lngCP = rRigheCT.RecSet.AbsolutePosition + 2
                    CodStampa = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODSTAMPA", ""))
                    CodArt = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODART", ""))
                    CodGruppo = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODGRUPPO", ""))
                    CodSottoGruppo = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODSOTTOGRUPPO", ""))
                    DscArt = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "DESCRIZIONE", ""))
                    DscEstesa = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "DESCRESTESA", ""))
                    dblQta = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "QTA", 0))
                    UM = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "UM", ""))
                    RicavoTot = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOTOT", 0))
                    
                    If CodStampa <> "" Then
                            DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodStampa
                    Else
                        If CodArt = "" Then
                            If CodSottoGruppo = "" Then
                                DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodGruppo
                            Else
                                DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodSottoGruppo
                            End If
                        Else
                            DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodArt
                        End If
                    End If
                    DocWord.tables(1).Cell(lngCP, 2).Range.Text = DscArt & " " & DscEstesa
                    dblTotImp = dblTotImp + RicavoTot
                    Call .dbSuccessivo(rRigheCT)
                    If Not .dbFineTab(rRigheCT) Then
                        Call DocWord.tables(1).Rows.Add
                    End If
                Loop
            End With
        
        Else
            strSql = "SELECT * FROM VISTA_GENERAPCC WHERE IDSESSIONE = :IDSESSIONE and RIFPROGRESSIVO = :RIFPROGRESSIVO ORDER BY NRRIGA"
            strSql = Replace(strSql, ":IDSESSIONE", hndDBArchivi.FormatoSQL(MXNU.IDSessione, DB_LONG))
            strSql = Replace(strSql, ":RIFPROGRESSIVO", hndDBArchivi.FormatoSQL(lngRifPr, DB_LONG))
            
            With MXDB
                Set rRigheCT = .dbCreaSS(hndDBArchivi, strSql)
                
                Do While Not .dbFineTab(rRigheCT)
                    lngCP = rRigheCT.RecSet.AbsolutePosition + 2
                    CodStampa = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODSTAMPA", ""))
                    CodArt = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODART", ""))
                    CodGruppo = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODGRUPPO", ""))
                    CodSottoGruppo = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "CODSOTTOGRUPPO", ""))
                    DscArt = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "DESCRIZIONE", ""))
                    DscEstesa = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "DESCRESTESA", ""))
                    Sottocommessa = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "SOTTOCOMMESSA", ""))
                    intExtra = .dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "EXTRA", 0)
                    strUnitaImm = .dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "UNITAIMM", "")
                    Lunghezza = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "LUNGHEZZA", ""))
                    Larghezza = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "LARGHEZZA", ""))
                    Altezza = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "ALTEZZA", ""))
                    Numero = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "NUMERO", ""))
                    dblQta = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "QTA", 0))
                    UM = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "UM", ""))
                    CostoUnit = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOUNITARIO", 0))
                    CostoUnitTot = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOUNITTOT", 0))
                    PercMaggC = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "PERCMAGGC", ""))
                    CostoTot = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOTOT", 0))
                    PercMaggR = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "PERCMAGGR", ""))
                    RicavoUnitTot = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOUNITTOT", 0))
                    RicavoTotLordo = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOTOTLORDO", 0))
                    PercSconto = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "PERCSCONTO", ""))
                    RicavoTotNetto = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOTOTNETTO", 0))
                    PercSicurezza = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "PERCSICUREZZA", ""))
                    RicavoSicurezza = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOSICUREZZA", 0))
                    RicavoTot = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOTOT", 0))
                    
                    
                    If CodStampa <> "" Then
                        ' se è stato impostato un codice di stampa viene stampato al posto del codart
                        DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodStampa
                    Else
                        If CodArt = "" Then
                            If CodSottoGruppo = "" Then
                                DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodGruppo
                            Else
                                DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodSottoGruppo
                            End If
                        Else
                            DocWord.tables(1).Cell(lngCP, 1).Range.Text = CodArt
                        End If
                    End If
                    DocWord.tables(1).Cell(lngCP, 2).Range.Text = DscArt & " " & DscEstesa
                    
                    'If bolRicavi And intTipoPrev <> 1 Then
                    If Left(strFil, 6) = "PCCPRC" Then
                        DocWord.tables(1).Cell(lngCP, 3).Range.Text = strUnitaImm
                        DocWord.tables(1).Cell(lngCP, 4).Range.Text = Sottocommessa
                        DocWord.tables(1).Cell(lngCP, 5).Range.Text = IIf(intExtra = 1, "SI", "")
                        
                        DocWord.tables(1).Cell(lngCP, 6).Range.Text = FrmtValore(Lunghezza, 0)
                        DocWord.tables(1).Cell(lngCP, 7).Range.Text = FrmtValore(Larghezza, 0)
                        DocWord.tables(1).Cell(lngCP, 8).Range.Text = FrmtValore(Altezza, 0)
                        DocWord.tables(1).Cell(lngCP, 9).Range.Text = FrmtValore(Numero, 0)
                        DocWord.tables(1).Cell(lngCP, 10).Range.Text = FrmtValore(dblQta, 0)
                        DocWord.tables(1).Cell(lngCP, 11).Range.Text = UM
                        DocWord.tables(1).Cell(lngCP, 12).Range.Text = FrmtValore(CostoUnit, 1)
                        DocWord.tables(1).Cell(lngCP, 13).Range.Text = PercMaggC
                        DocWord.tables(1).Cell(lngCP, 14).Range.Text = FrmtValore(CostoUnitTot, 1)
                        DocWord.tables(1).Cell(lngCP, 15).Range.Text = FrmtValore(CostoTot, 2)
                        DocWord.tables(1).Cell(lngCP, 16).Range.Text = FrmtValore(RicavoUnitTot, 1)
                        DocWord.tables(1).Cell(lngCP, 17).Range.Text = PercMaggR
                        DocWord.tables(1).Cell(lngCP, 18).Range.Text = FrmtValore(RicavoTotLordo, 2)
                        DocWord.tables(1).Cell(lngCP, 19).Range.Text = PercSconto
                        'DocWord.tables(1).Cell(lngCP, 20).Range.Text = FrmtValore(RicavoTotNetto, 2)
                        DocWord.tables(1).Cell(lngCP, 20).Range.Text = PercSicurezza
                        DocWord.tables(1).Cell(lngCP, 21).Range.Text = FrmtValore(RicavoSicurezza, 2)
                        DocWord.tables(1).Cell(lngCP, 22).Range.Text = FrmtValore(RicavoTot, 2)
                    Else
                        DocWord.tables(1).Cell(lngCP, 3).Range.Text = IIf(intExtra = 1, "SI", "") 'Sottocommessa
                        DocWord.tables(1).Cell(lngCP, 4).Range.Text = FrmtValore(Lunghezza, 0)
                        DocWord.tables(1).Cell(lngCP, 5).Range.Text = FrmtValore(Larghezza, 0)
                        DocWord.tables(1).Cell(lngCP, 6).Range.Text = FrmtValore(Altezza, 0)
                        DocWord.tables(1).Cell(lngCP, 7).Range.Text = FrmtValore(Numero, 0)
                        DocWord.tables(1).Cell(lngCP, 8).Range.Text = FrmtValore(dblQta, 0)
                        DocWord.tables(1).Cell(lngCP, 9).Range.Text = UM
                        If intTipoPrev = 1 Then
                            dblImp1 = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOUNITTOT", 0))
                        Else
                            DocWord.tables(1).Cell(1, 10).Range.Text = "Costo Unit."
                            dblImp1 = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOUNITTOT", 0))
                        End If
    
                        If dblImp1 = "0" Then
                            dblImp1 = ""
                        Else
                            dblImp1 = mxutil.FormatCampo(dblImp1, MXNU.FORMATO_EURO_UNITARIO)
                        End If
                        DocWord.tables(1).Cell(lngCP, 10).Range.Text = dblImp1 'IIf(CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOUNITTOT", "")) = "0", "", CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOUNITTOT", "")))
    
                        If intTipoPrev = 1 Then
                            dblImp2 = CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOTOT", 0))
                        Else
                            DocWord.tables(1).Cell(1, 11).Range.Text = "Costo Tot."
                            dblImp2 = CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOTOT", 0))
                        End If
                        
                        If dblImp2 = "0" Then
                            dblImp2 = ""
                        Else
                            dblImp2 = mxutil.FormatCampo(dblImp2, MXNU.FORMATO_EURO_TOTALE)
                        End If
                        DocWord.tables(1).Cell(lngCP, 11).Range.Text = dblImp2 'IIf(CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOTOT", "")) = "0", "", CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOTOT", "")))
                    End If


                    dblTotRic = dblTotRic + CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOTOT", ""))
                    

                    
                    If Left(strFil, 6) = "PCCPRC" Then
                        dblTotImp = dblTotImp + CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOTOT", ""))
                    Else
                        If intTipoPrev = 1 Then
                            dblTotImp = dblTotImp + CDbl(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "RICAVOTOT", ""))
                        Else
                            dblTotImp = dblTotImp + CStr(.dbGetCampo(rRigheCT, TIPO_SNAPSHOT, "COSTOTOT", ""))
                        End If
                    End If

                    Call .dbSuccessivo(rRigheCT)
                    If Not .dbFineTab(rRigheCT) Then
                        Call DocWord.tables(1).Rows.Add
                    End If
                Loop
            End With
        End If
        
        strSql = "SELECT DIVISA FROM TABCAMBI WHERE CODICE = :CODICE"
        strSql = Replace(strSql, ":CODICE", hndDBArchivi.FormatoSQL(MXDB.dbGetCampo(hSSSrc, TIPO_SNAPSHOT, "CODCAMBIO", 0), DB_INTEGER))
        Set rSql = MXDB.dbCreaSS(hndDBArchivi, strSql, TIPO_SNAPSHOT)
        
        
        ' Assegna dato al campo totale del modulo word
        If DocWord.Bookmarks.Exists("TotImp") Then
            DocWord.Bookmarks("TotImp").Range.Text = "Totale " & MXDB.dbGetCampo(rSql, TIPO_SNAPSHOT, "DIVISA", "") & " " & mxutil.FormatCampo(dblTotImp, MXNU.FORMATO_EURO_TOTALE)
        End If
        ' Assegna dato al campo totale del modulo word
        If DocWord.Bookmarks.Exists("TotRic") Then
            DocWord.Bookmarks("TotRic").Range.Text = "Totale " & MXDB.dbGetCampo(rSql, TIPO_SNAPSHOT, "DIVISA", "") & " " & mxutil.FormatCampo(dblTotRic, MXNU.FORMATO_EURO_TOTALE)
        End If
        Call MXDB.dbChiudiSS(rSql)
        Set rSql = Nothing

    Call MXDB.dbChiudiSS(hSSSrc)
    Set hSSSrc = Nothing

    Dim strNome As String
    Dim strPath As String
    Call mxutil.scomponiPathFile(strNomeDOT, strPath, strNome)
    
    If Not DocWord Is Nothing Then
        ' Stampa documento
        'strNomeFile = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\PREVENTIVI\" & strTipoPrev & " " & lngNrPrev & "_" & strBis & ".doc" ' " _" & CStr(Date) & ".doc"
        
        ' modifica del 18/04/2003
        If strPathWORD = "" Or IsNull(strPathWORD) Then
            strPathWORD = MXNU.PercorsoPers
            Call MXNU.MsgBoxEX("Impossibile trovare il percorso in cui salvare il file Word, il file verrà salvato in : " & strPathWORD, vbInformation, "Salvataggio Preventivi Word")
        End If
        
        strNomeFile = strPathWORD & "\" & strTipoPrev & " " & lngNrPrev & "_" & strBis & ".doc" ' " _" & CStr(Date) & ".doc"
        DocWord.SaveAs strNomeFile
        mNomeFile = strNomeFile

        
        If (MXNU.MsgBoxEX("Confermi la stampa del preventivo?", vbYesNo + vbQuestion, 1007) = vbYes) Then

            DocWord.PrintOut
            For I = 1 To 10000
            Next
            MsgBox "Premere Ok al completamento della stampa"

        End If
        DocWord.Close
        Set DocWord = Nothing
        AppWord.QUIT
        Set AppWord = Nothing
    End If

' ObjInserter
'---------------------------------------------
    'Resume Next
    Exit Function

ComProcedi_Click_ErrHandler:
    If Err.Number <> 462 And Err.Number <> 190 And Err.Number <> 5941 Then
    Dim cLogString As String
    cLogString = "Error in StampaPrev: Sub ComProcedi_Click (" & Erl & ")" & vbCrLf & Err.Description & vbCr & vbCr & "(Error #" & Err.Number & ")"
    MsgBox cLogString, vbCritical + vbOKOnly, "Errore!"
    End If
'---------------------------------------------
Resume Next

End Function

Function CercaContratto(ByVal IDTESTA As Double) As Boolean
    Dim strSql As String
    Dim rSql As CRecordSet
    
    strSql = "SELECT TIPODOC, NUMERODOC, BIS, MAX(CODDEPOSITO) AS COMMESSA FROM RIGHEDOCUMENTI "
    strSql = strSql & " WHERE IDTESTA = :IDTESTA"
    strSql = strSql & " GROUP BY TIPODOC, NUMERODOC, BIS"
    strSql = Replace(strSql, ":IDTESTA", hndDBArchivi.FormatoSQL(IDTESTA, DB_DECIMAL))
    
    With MXDB
        Set rSql = .dbCreaSS(hndDBArchivi, strSql)
        If Not .dbFineTab(rSql) Then
            mTipoDoc = .dbGetCampo(rSql, TIPO_SNAPSHOT, "TIPODOC", "")
            mNumeroDoc = .dbGetCampo(rSql, TIPO_SNAPSHOT, "NUMERODOC", 0)
            mBis = .dbGetCampo(rSql, TIPO_SNAPSHOT, "BIS", "")
            mCommessa = .dbGetCampo(rSql, TIPO_SNAPSHOT, "COMMESSA", "")
            CercaContratto = True
        Else
            CercaContratto = False
        End If
        Call .dbChiudiSS(rSql)
        Set rSql = Nothing
    End With
    
End Function

Function ProgCommessa(ByVal strCodice As String) As Boolean
    Dim strSql As String
    Dim rSql As CRecordSet
    
    strSql = "SELECT PROGRESSIVO FROM TESTECOMMESSECLI "
    strSql = strSql & " WHERE CODICE = :CODICE"
    strSql = Replace(strSql, ":CODICE", hndDBArchivi.FormatoSQL(strCodice, DB_TEXT))
    
    With MXDB
        Set rSql = .dbCreaSS(hndDBArchivi, strSql)
        If Not .dbFineTab(rSql) Then
            mProgCommessa = .dbGetCampo(rSql, TIPO_SNAPSHOT, "PROGRESSIVO", "")
            ProgCommessa = True
        Else
            ProgCommessa = False
        End If
        Call .dbChiudiSS(rSql)
        Set rSql = Nothing
    End With
    
End Function


Function SalvaSpread(ByVal objProfilo As Variant) As Boolean
    ' funzione che salva il profilo
    If objProfilo.ttabmodificata Then
    
    End If
End Function
