VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDispArticolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarGiacenza As Variant
Private mvarDisponibilita As Variant
Private mvarCarichi As Variant
Private mvarScarichi As Variant
Private mvarResiDaCarico As Variant
Private mvarResidaScarico As Variant
Private mvarOrdinato As Variant
Private mvarImpegnato As Variant
Private mvarGiacenzaIniziale As Variant
'gestione riservazioni
Private mDecGiacenzaLibera As Variant
Private mDecGiacenzaRiservata As Variant
Private mDecGiacezaDisponibile As Variant

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'                           proprietà pubbliche della classe
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Property Let GiacenzaIniziale(ByVal vData As Variant)
    mvarGiacenzaIniziale = vData
End Property

Public Property Set GiacenzaIniziale(ByVal vData As Variant)
    Set mvarGiacenzaIniziale = vData
End Property

Public Property Get GiacenzaIniziale() As Variant
    GiacenzaIniziale = mvarGiacenzaIniziale
End Property

Public Property Let Impegnato(ByVal vData As Variant)
    mvarImpegnato = vData
End Property

Public Property Set Impegnato(ByVal vData As Variant)
    Set mvarImpegnato = vData
End Property

Public Property Get Impegnato() As Variant
    Impegnato = mvarImpegnato
End Property

Public Property Let Ordinato(ByVal vData As Variant)
    mvarOrdinato = vData
End Property

Public Property Set Ordinato(ByVal vData As Variant)
    Set mvarOrdinato = vData
End Property

Public Property Get Ordinato() As Variant
    Ordinato = mvarOrdinato
End Property

Public Property Let ResidaScarico(ByVal vData As Variant)
    mvarResidaScarico = vData
End Property

Public Property Set ResidaScarico(ByVal vData As Variant)
    Set mvarResidaScarico = vData
End Property

Public Property Get ResidaScarico() As Variant
    ResidaScarico = mvarResidaScarico
End Property

Public Property Let ResiDaCarico(ByVal vData As Variant)
    mvarResiDaCarico = vData
End Property

Public Property Set ResiDaCarico(ByVal vData As Variant)
    Set mvarResiDaCarico = vData
End Property

Public Property Get ResiDaCarico() As Variant
    ResiDaCarico = mvarResiDaCarico
End Property

Public Property Let Scarichi(ByVal vData As Variant)
    mvarScarichi = vData
End Property

Public Property Set Scarichi(ByVal vData As Variant)
    Set mvarScarichi = vData
End Property

Public Property Get Scarichi() As Variant
    Scarichi = mvarScarichi
End Property

Public Property Let Carichi(ByVal vData As Variant)
    mvarCarichi = vData
End Property

Public Property Set Carichi(ByVal vData As Variant)
    Set mvarCarichi = vData
End Property

Public Property Get Carichi() As Variant
    Carichi = mvarCarichi
End Property

Public Property Let Disponibilita(ByVal vData As Variant)
    mvarDisponibilita = vData
End Property

Public Property Set Disponibilita(ByVal vData As Variant)
    Set mvarDisponibilita = vData
End Property

Public Property Get Disponibilita() As Variant
    Disponibilita = mvarDisponibilita
End Property

Public Property Let Giacenza(ByVal vData As Variant)
    mvarGiacenza = vData
End Property

Public Property Set Giacenza(ByVal vData As Variant)
    Set mvarGiacenza = vData
End Property

Public Property Get Giacenza() As Variant
    Giacenza = mvarGiacenza
End Property

Public Property Let GiacenzaLibera(ByVal vData As Variant)
    mDecGiacenzaLibera = vData
End Property

Public Property Get GiacenzaLibera() As Variant
    GiacenzaLibera = mDecGiacenzaLibera
End Property

Public Property Let GiacenzaRiservata(ByVal vData As Variant)
    mDecGiacenzaRiservata = vData
End Property

Public Property Get GiacenzaRiservata() As Variant
    GiacenzaRiservata = mDecGiacenzaRiservata
End Property

Public Property Let GiacezaDisponibile(ByVal vData As Variant)
    mDecGiacezaDisponibile = vData
End Property

Public Property Get GiacezaDisponibile() As Variant
    GiacezaDisponibile = mDecGiacezaDisponibile
End Property

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'                           funzioni pubbliche della classe
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Function CalcolaDisp(ByVal strCodArticolo As String, _
                            ByVal intEsercizio As Integer, _
                            ByVal bolCalcola2UM As Boolean, _
                            ByVal strCodDeposito As String, _
                            ByVal strCodUbicazione As String, _
                            ByVal strPartita As String, _
                            Optional ByVal bolListaDepositi As Boolean = False, _
                            Optional ByVal bolListaUbicazioni As Boolean = False, _
                            Optional ByVal DaData As Variant, _
                            Optional ByVal AData As Variant)
    
    Dim hSS As MXKit.CRecordSet
    Dim intq As Integer
    Dim strWHE As String
    Dim strWHEIniz As String
    Dim strSQL As String
    Dim strGROUP_BY As String
    
    strWHE = "CodArt=" & hndDBArchivi.FormatoSQL(strCodArticolo, DB_TEXT)
    strGROUP_BY = "CodArt"
    If (strCodDeposito <> "") Then
        If (bolListaDepositi) Then
            strWHE = strWHE & " AND CodDeposito IN (" & strCodDeposito & ")"
        Else
            strWHE = strWHE & " AND CodDeposito=" & hndDBArchivi.FormatoSQL(strCodDeposito, DB_TEXT)
        End If
        strGROUP_BY = strGROUP_BY & ",CodDeposito"
    End If
    If (strCodUbicazione <> MAG_TUTTE_LE_UBICAZIONI) Then
        If (bolListaUbicazioni) Then
            strWHE = strWHE & " AND CodUbicazione IN (" & strCodUbicazione & ")"
        Else
            strWHE = strWHE & " AND CodUbicazione=" & hndDBArchivi.FormatoSQL(strCodUbicazione, DB_TEXT)
        End If
        strGROUP_BY = strGROUP_BY & ",CodUbicazione"
    End If
    
    If (strPartita <> MAG_TUTTE_LE_PARTITE) Then
        strWHE = strWHE & " AND NRifPartita=" & hndDBArchivi.FormatoSQL(strPartita, DB_TEXT)
        strGROUP_BY = strGROUP_BY & ",NRifPartita"
    End If
    strWHEIniz = strWHE
    If Not (IsMissing(DaData)) Then
        strWHE = strWHE & " AND DataMov >=" & hndDBArchivi.FormatoSQL(DaData, DB_DATE)
        strWHEIniz = strWHEIniz & " AND DataMov <" & hndDBArchivi.FormatoSQL(DaData, DB_DATE)
    Else
        strWHEIniz = strWHEIniz & " AND Esercizio<=" & (intEsercizio - 1)
    End If
    If Not (IsMissing(AData)) Then
        strWHE = strWHE & " AND DataMov <=" & hndDBArchivi.FormatoSQL(AData, DB_DATE)
    End If
    
    If Not bolCalcola2UM Then
        strSQL = "SELECT SUM(Carico+ResoDaScarico-Scarico-ResoDaCarico) AS Disp, SUM(Ordinato) AS Ordinato, SUM(Impegnato) AS Impegnato FROM VistaGiacenze WHERE " & strWHEIniz & " GROUP BY " & strGROUP_BY
    Else
        strSQL = "SELECT SUM(Carico2UM+ResoDaScarico2UM-Scarico2UM-ResoDaCarico2UM) AS Disp, SUM(Ordinato2UM) AS Ordinato, SUM(Impegnato2UM) AS Impegnato FROM VistaGiacenze WHERE " & strWHEIniz & " GROUP BY " & strGROUP_BY
    End If
        
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL)
    intq = Not MXDB.dbFineTab(hSS)
    mvarGiacenzaIniziale = 0
    mvarOrdinato = 0
    mvarImpegnato = 0
    While intq
        mvarGiacenzaIniziale = CDec(mvarGiacenzaIniziale) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Disp", 0))
        mvarOrdinato = CDec(mvarOrdinato) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Ordinato", 0))
        mvarImpegnato = CDec(mvarImpegnato) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Impegnato", 0))
        
        intq = MXDB.dbSuccessivo(hSS)
    Wend
    intq = MXDB.dbChiudiSS(hSS)
    
    If Not bolCalcola2UM Then
        strSQL = "SELECT SUM(Carico) AS Carichi,SUM(ResoDaScarico) AS ResiDaScarico,SUM(Scarico) AS Scarichi,SUM(ResoDaCarico) AS ResiDaCarico,SUM(Ordinato) AS Ordinato,SUM(Impegnato) AS Impegnato FROM VistaGiacenze WHERE " & strWHE & " AND Esercizio=" & intEsercizio & " GROUP BY " & strGROUP_BY
    Else
        strSQL = "SELECT SUM(Carico2UM) AS Carichi,SUM(ResoDaScarico2UM) AS ResiDaScarico,SUM(Scarico2UM) AS Scarichi,SUM(ResoDaCarico2UM) AS ResiDaCarico,SUM(Ordinato2UM) AS Ordinato,SUM(Impegnato2UM) AS Impegnato FROM VistaGiacenze WHERE " & strWHE & " AND Esercizio=" & intEsercizio & " GROUP BY " & strGROUP_BY
    End If
    mvarCarichi = 0
    mvarScarichi = 0
    mvarResiDaCarico = 0
    mvarResidaScarico = 0
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL)
    intq = Not MXDB.dbFineTab(hSS)
    While intq
        mvarCarichi = CDec(mvarCarichi) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Carichi", 0))
        mvarScarichi = CDec(mvarScarichi) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Scarichi", 0))
        mvarResiDaCarico = CDec(mvarResiDaCarico) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "ResiDaCarico", 0))
        mvarResidaScarico = CDec(mvarResidaScarico) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "ResiDaScarico", 0))
        mvarOrdinato = CDec(mvarOrdinato) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Ordinato", 0))
        mvarImpegnato = CDec(mvarImpegnato) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Impegnato", 0))
        
        intq = MXDB.dbSuccessivo(hSS)
    Wend
    mvarGiacenza = mvarGiacenzaIniziale + mvarCarichi + mvarResidaScarico - mvarScarichi - mvarResiDaCarico
    mvarDisponibilita = mvarGiacenza + mvarOrdinato - mvarImpegnato
    intq = MXDB.dbChiudiSS(hSS)
    
End Function
Public Sub CalcolaDispTot(ByVal strCodArticolo As String, _
                            ByVal intEsercizio As Integer, _
                            ByVal bolCalcola2UM As Boolean, _
                            ByVal strCodDeposito As String, _
                            ByVal strCodUbicazione As String, _
                            ByVal strPartita As String, _
                            Optional ByVal bolListaDepositi As Boolean = False, _
                            Optional ByVal bolListaUbicazioni As Boolean = False)
    
    Dim hSS As MXKit.CRecordSet
    Dim intq As Integer
    Dim strWHE As String
    Dim strSQL As String
    Dim strGROUP_BY As String
    Dim intEsercizioGiac As Integer
    
    strWHE = "CodArt=" & hndDBArchivi.FormatoSQL(strCodArticolo, DB_TEXT)
    strGROUP_BY = "Esercizio,CodArt"
    If (strCodDeposito <> "") Then
        If (bolListaDepositi) Then
            strWHE = strWHE & " AND CodDeposito IN (" & strCodDeposito & ")"
        Else
            strWHE = strWHE & " AND CodDeposito=" & hndDBArchivi.FormatoSQL(strCodDeposito, DB_TEXT)
        End If
        strGROUP_BY = strGROUP_BY & ",CodDeposito"
    End If
    If (strCodUbicazione <> MAG_TUTTE_LE_UBICAZIONI) Then
        If (bolListaUbicazioni) Then
            strWHE = strWHE & " AND CodUbicazione IN (" & strCodUbicazione & ")"
        Else
            strWHE = strWHE & " AND CodUbicazione=" & hndDBArchivi.FormatoSQL(strCodUbicazione, DB_TEXT)
        End If
        strGROUP_BY = strGROUP_BY & ",CodUbicazione"
    End If
    
    If (strPartita <> MAG_TUTTE_LE_PARTITE) Then
        strWHE = strWHE & " AND NRifPartita=" & hndDBArchivi.FormatoSQL(strPartita, DB_TEXT)
        strGROUP_BY = strGROUP_BY & ",NRifPartita"
    End If
    strWHE = strWHE & " AND Esercizio<=" & intEsercizio
    
    If Not bolCalcola2UM Then
        strSQL = "SELECT Esercizio, SUM(Carico) AS Carichi,SUM(ResoDaScarico) AS ResiDaScarico,SUM(Scarico) AS Scarichi,SUM(ResoDaCarico) AS ResiDaCarico,SUM(Ordinato) AS Ordinato,SUM(Impegnato) AS Impegnato FROM VistaGiacenzeTot WHERE " & strWHE & " GROUP BY " & strGROUP_BY & " ORDER BY Esercizio"
    Else
        strSQL = "SELECT Esercizio, SUM(Carico2UM) AS Carichi,SUM(ResoDaScarico2UM) AS ResiDaScarico,SUM(Scarico2UM) AS Scarichi,SUM(ResoDaCarico2UM) AS ResiDaCarico,SUM(Ordinato2UM) AS Ordinato,SUM(Impegnato2UM) AS Impegnato FROM VistaGiacenzeTot WHERE " & strWHE & " GROUP BY " & strGROUP_BY & " ORDER BY Esercizio"
    End If
        
    Call MXDB.EnabledVisteIndicizzate(hndDBArchivi, True)
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL)
    intq = Not MXDB.dbFineTab(hSS)
    mvarGiacenzaIniziale = 0
    mvarOrdinato = 0
    mvarImpegnato = 0
    mvarCarichi = 0
    mvarScarichi = 0
    mvarResiDaCarico = 0
    mvarResidaScarico = 0
    While intq
        intEsercizioGiac = MXDB.dbGetCampo(hSS, NO_REPOSITION, "Esercizio", 0)
        If intEsercizioGiac < intEsercizio Then
            mvarGiacenzaIniziale = CDec(mvarGiacenzaIniziale) + _
                                   CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Carichi", 0)) + _
                                   CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "ResiDaScarico", 0)) - _
                                   CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Scarichi", 0)) - _
                                   CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "ResiDaCarico", 0))
            mvarOrdinato = CDec(mvarOrdinato) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Ordinato", 0))
            mvarImpegnato = CDec(mvarImpegnato) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Impegnato", 0))
        Else
            mvarCarichi = CDec(mvarCarichi) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Carichi", 0))
            mvarScarichi = CDec(mvarScarichi) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Scarichi", 0))
            mvarResiDaCarico = CDec(mvarResiDaCarico) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "ResiDaCarico", 0))
            mvarResidaScarico = CDec(mvarResidaScarico) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "ResiDaScarico", 0))
            mvarOrdinato = CDec(mvarOrdinato) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Ordinato", 0))
            mvarImpegnato = CDec(mvarImpegnato) + CDec(MXDB.dbGetCampo(hSS, NO_REPOSITION, "Impegnato", 0))
        End If
        
        intq = MXDB.dbSuccessivo(hSS)
    Wend
    intq = MXDB.dbChiudiSS(hSS)
    Call MXDB.EnabledVisteIndicizzate(hndDBArchivi, False)
    
    mvarGiacenza = mvarGiacenzaIniziale + mvarCarichi + mvarResidaScarico - mvarScarichi - mvarResiDaCarico
    mvarDisponibilita = mvarGiacenza + mvarOrdinato - mvarImpegnato
    
End Sub

'------------------------------------------------------------
'nome:          CalcolaGiacenzaRiservata
'descrizione:   calcola la giacenza riservata in base al riferimento commessa
'parametri:
'ritorno:
'annotazioni:
'------------------------------------------------------------
Public Function CalcolaGiacenzaRiservata(ByVal strCodArticolo As String, _
                                        ByVal strRifCommCli As String, _
                                        Optional ByVal bolCalcolaDisp As Boolean = False) As Boolean
Dim bolRes As Boolean
Dim strSQL As String
Dim decGiacRiservate As Variant
Dim hSS As MXKit.CRecordSet

    bolRes = True
    On Local Error GoTo CalcolaGiacenzaRiservata_ERR
    'ricalcolo giacenze
    If bolCalcolaDisp Then
        Call CalcolaDisp(strCodArticolo, MXNU.AnnoAttivo, False, "", "", "")
    End If
    'calcolo giacenza libera (=giacenza totale - somma giacenze riservate)
    strSQL = "select sum(QTAUMBASE) QTARIS" _
        & " from TABLEGAMIDISP" _
        & " where CODART=" & hndDBArchivi.FormatoSQL(strCodArticolo, DB_TEXT) _
        & " and RIFCOMMCLI<>''"
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL)
    decGiacRiservate = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "QTARIS", CDec(0))
    MXDB.dbChiudiSS hSS
    mDecGiacenzaLibera = mvarGiacenza - decGiacRiservate
    If mDecGiacenzaLibera < CDec(0) Then mDecGiacenzaLibera = CDec(0)
    'calcolo  giacenza riservata alla commessa
    If Len(strRifCommCli) = 0 Then
        '   se commessa blank = giacenza libera
        mDecGiacenzaRiservata = mDecGiacenzaLibera
    Else
        '   se commessa non blank = giacenza riservata alla commessa
        strSQL = "select sum(QTAUMBASE) QTARIS" _
            & " from TABLEGAMIDISP" _
            & " where CODART=" & hndDBArchivi.FormatoSQL(strCodArticolo, DB_TEXT) _
            & " and RIFCOMMCLI=" & hndDBArchivi.FormatoSQL(strRifCommCli, DB_TEXT)
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL)
        mDecGiacenzaRiservata = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "QTARIS", CDec(0))
        MXDB.dbChiudiSS hSS
    End If
    'calcolo giacenza disponibile alla commessa
    If Len(strRifCommCli) = 0 Then
        '   se commessa blank = giacenza libera
        mDecGiacezaDisponibile = mDecGiacenzaLibera
    Else
        '   se commessa non blank = giacenza libera + giacenza riservata
        mDecGiacezaDisponibile = mDecGiacenzaLibera + mDecGiacenzaRiservata
    End If
    
CalcolaGiacenzaRiservata_END:
    CalcolaGiacenzaRiservata = bolRes
    On Local Error GoTo 0
    Exit Function

CalcolaGiacenzaRiservata_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("CalcolaGiacenzaRiservata", lngErrCod, strErrDsc))
    Resume CalcolaGiacenzaRiservata_END
End Function

