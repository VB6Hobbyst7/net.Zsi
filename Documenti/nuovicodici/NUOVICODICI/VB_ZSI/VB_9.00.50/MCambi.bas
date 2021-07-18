Attribute VB_Name = "MCambi"
Option Explicit
DefLng A-Z

Private MCambioEuro As Variant
Public Enum setTipoDecimali
    TDEC_IMPUNITARIO = -1
    TDEC_IMPTOTALE = -2
End Enum
    
'NOME           : LeggiValoreCambio
'DESCRIZIONE    : legge il valore cambio per la divisa/data specificate
'PARAMETRO 1    : codice divisa
'PARAMETRO 2    : data cambio
'PARAMETRO 3    : visualizza messaggio errore si/no
'RITORNO        : restituisce il valore cambio
Function LeggiValoreCambio(ByVal intDivisa As Integer, _
                            ByVal vntData As Variant, _
                            Optional bolMessaggi As Boolean = True, Optional bolMessaggioMostrato As Boolean = False) As Variant
    
    bolMessaggioMostrato = False
    If (intDivisa = 0) Or (Not IsDate(vntData)) Then
        LeggiValoreCambio = 1
    Else
        Dim strSQL As String
        Dim hSS As CRecordset
        
        If vntData >= MXNU.DataInizioTriangolazione And MonetaEURO(intDivisa) Then
            LeggiValoreCambio = 1
        Else
            strSQL = "SELECT Valore" _
                    & " FROM ValoriCambio" _
                    & " WHERE (Data = " & hndDBArchivi.FormatoSQL(vntData, DB_DATE) & ")" _
                    & " AND (CodCambio = " & hndDBArchivi.FormatoSQL(intDivisa, DB_INTEGER) & ")"
            Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
            If (MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
                If bolMessaggi Then
                    Call MXNU.MsgBoxEX(1874, vbOKOnly + vbExclamation, 1007, Array(CStr(vntData)))
                    bolMessaggioMostrato = True
                End If
                LeggiValoreCambio = 1
            Else
                LeggiValoreCambio = fdec(CDec(MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Valore", 1)), 6)
            End If
            Call MXDB.dbChiudiSS(hSS)
       End If
    End If
End Function


Function ConvertiDivisa(ByVal DivisaIn As Long, ByVal DivisaOut As Long, Importo As Variant, TipoDecimali As setTipoDecimali, vntDataOp As Date, Optional ValoreCambio As Variant = 1) As Variant
    Dim vntValore As Variant, intNDec As Integer
    If DivisaIn <> DivisaOut Then
        If (MonetaEURO(DivisaIn) Or MonetaEURO(DivisaOut)) And (vntDataOp >= MXNU.DataInizioTriangolazione) Then
            vntValore = Divisa2EURO(DivisaIn, Importo, 6, vntDataOp, ValoreCambio)
            vntValore = EURO2Divisa(DivisaOut, vntValore, 6, vntDataOp, ValoreCambio)
            intNDec = LeggiNDecimali(DivisaOut, TipoDecimali)
            vntValore = CDec(fdec(vntValore, intNDec))
        Else
            If ValoreCambio = 0 Then ValoreCambio = 1
            intNDec = LeggiNDecimali(DivisaOut, TipoDecimali)
            vntValore = CDec(fdec((Importo / ValoreCambio), intNDec))
        End If
    Else
        vntValore = Importo
    End If
    ConvertiDivisa = vntValore
End Function

'Converte la divisa (di Tipo Euro) passata, al corrispondente valore espresso in EURO secondo
'il valore fisso valuta/Euro definito nella tabella cambi
Public Function Divisa2EURO(ByVal CodDivisa As Long, ByVal valore As Variant, ByVal TipoDecimali As setTipoDecimali, ByVal vntDataOp As Variant, Optional ByVal ValoreCambio As Variant = 1) As Variant
    Static intNDecimali As Integer
    Static vntCambioEuro As Variant
    Static LastCodDivisa As Long
    Static LastTipoDecimali As Long
    Static bolLeggiCambioEuro As Boolean
    Static bolNotInitVar As Boolean
    Static DataOp As Date
    Static UltDitta As String
    
    If Not IsDate(vntDataOp) Then Call MXNU.MsgBoxEX("Data di valorizzazione non valida!", vbCritical, "Divisa2EURO")
    If Not bolNotInitVar Then LastCodDivisa = -1: bolNotInitVar = True
        
    If CodDivisa <> LastCodDivisa Or TipoDecimali <> LastTipoDecimali Or DataOp <> vntDataOp Or MXNU.DittaAttiva & MXNU.AnnoAttivo <> UltDitta Then
        If Not MonetaEURO(CodDivisa) Then
            vntCambioEuro = ValoreCambio
            If vntCambioEuro = 0 Or vntCambioEuro = "" Then vntCambioEuro = 1
            bolLeggiCambioEuro = False
        Else
            bolLeggiCambioEuro = True
        End If
        If bolLeggiCambioEuro Then
            If vntDataOp < MXNU.DataInizioTriangolazione Then
                vntCambioEuro = ValoreCambio
                If vntCambioEuro = 1 Then
                    vntCambioEuro = LeggiValoreCambio(CodDivisa, vntDataOp, False)
                End If
                If vntCambioEuro = 1 Then vntCambioEuro = MCambioEuro
            Else
                vntCambioEuro = MCambioEuro
            End If
        End If
        intNDecimali = LeggiNDecimali(MXNU.CodCambioEuro, TipoDecimali)
        LastCodDivisa = CodDivisa
        LastTipoDecimali = TipoDecimali
        UltDitta = MXNU.DittaAttiva & MXNU.AnnoAttivo
        If IsDate(vntDataOp) Then DataOp = vntDataOp
    Else
        If Not bolLeggiCambioEuro Then vntCambioEuro = ValoreCambio
        If vntCambioEuro = "" Then vntCambioEuro = 1
    End If
    Divisa2EURO = CDec(fdec((valore / vntCambioEuro), intNDecimali))
    
End Function


'Converte un valore espresso in EURO nel corrispondente valore in valuta secondo
'il valore fisso valuta/Euro definito nella tabella cambi
Public Function EURO2Divisa(ByVal CodDivisa As Long, ByVal valore As Variant, ByVal TipoDecimali As setTipoDecimali, ByVal vntDataOp As Variant, Optional ByVal ValoreCambio As Variant = 1) As Variant
    Static vntCambioEuro As Variant
    Static intNDecimali As Integer
    Static bolLeggiCambioEuro As Boolean
    Static LastCodDivisa As Long
    Static LastTipoDecimali As Long
    Static bolNotInitVar As Boolean
    Static DataOp As Date
    Static UltDitta As String
    
    If Not IsDate(vntDataOp) Then Call MXNU.MsgBoxEX("Data di valorizzazione non valida!", vbCritical, "Euro2Divisa")
    If Not bolNotInitVar Then LastCodDivisa = -1: bolNotInitVar = True
        
    If CodDivisa <> LastCodDivisa Or TipoDecimali <> LastTipoDecimali Or DataOp <> vntDataOp Or MXNU.DittaAttiva & MXNU.AnnoAttivo <> UltDitta Then
        If Not MonetaEURO(CodDivisa) Then
            vntCambioEuro = ValoreCambio
            If vntCambioEuro = 0 Or vntCambioEuro = "" Then vntCambioEuro = 1
            bolLeggiCambioEuro = False
        Else
            bolLeggiCambioEuro = True
        End If
        If bolLeggiCambioEuro Then
            If vntDataOp < MXNU.DataInizioTriangolazione Then
                vntCambioEuro = ValoreCambio
                If vntCambioEuro = 1 Then
                    vntCambioEuro = LeggiValoreCambio(CodDivisa, vntDataOp, False)
                End If
                If vntCambioEuro = 1 Then vntCambioEuro = MCambioEuro
            Else
                vntCambioEuro = MCambioEuro
            End If
        End If
        
        intNDecimali = LeggiNDecimali(CodDivisa, TipoDecimali)
        LastCodDivisa = CodDivisa
        LastTipoDecimali = TipoDecimali
        UltDitta = MXNU.DittaAttiva & MXNU.AnnoAttivo
        If IsDate(vntDataOp) Then DataOp = vntDataOp
    Else
        If Not bolLeggiCambioEuro Then vntCambioEuro = ValoreCambio
        If vntCambioEuro = "" Then vntCambioEuro = 1
    End If
    EURO2Divisa = CDec(fdec((valore * vntCambioEuro), intNDecimali))
    
End Function



Public Function LeggiNDecimali(ByVal CodDivisa As Long, TipoDecimali As setTipoDecimali, Optional ImportoPerDecimali As Variant) As Integer
    Dim intNDec As Variant
    If IsMissing(ImportoPerDecimali) Then
        Dim q%, hSS As CRecordset, strNomeCmpDec$
        If CodDivisa = 0 Then
            If TipoDecimali = TDEC_IMPTOTALE Then
                LeggiNDecimali = MXNU.DecimaliLireTotale
            ElseIf TipoDecimali = TDEC_IMPUNITARIO Then
                LeggiNDecimali = MXNU.DecimaliLireUnitario
            Else
                LeggiNDecimali = TipoDecimali
            End If
            Exit Function
        ElseIf CodDivisa = MXNU.CodCambioEuro Then
            If TipoDecimali = TDEC_IMPTOTALE Then
                LeggiNDecimali = MXNU.DecimaliEuroTotale
            ElseIf TipoDecimali = TDEC_IMPUNITARIO Then
                LeggiNDecimali = MXNU.DecimaliEuroUnitario
            Else
                LeggiNDecimali = TipoDecimali
            End If
            Exit Function
        Else
            If TipoDecimali = TDEC_IMPTOTALE Then
                strNomeCmpDec = "NDecimaliTotale"
            ElseIf TipoDecimali = TDEC_IMPUNITARIO Then
                strNomeCmpDec = "NDecimaliUnitario"
            Else
                LeggiNDecimali = TipoDecimali
                Exit Function
            End If
        End If
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT " & strNomeCmpDec & " FROM TabCambi WHERE Codice=" & CodDivisa)
        intNDec = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, strNomeCmpDec, 0)
        q = MXDB.dbChiudiSS(hSS)
        LeggiNDecimali = intNDec
    Else
        If ImportoPerDecimali < 10 Then
            intNDec = 5
        ElseIf ImportoPerDecimali < 100 Then
            intNDec = 4
        ElseIf ImportoPerDecimali < 1000 Then
            intNDec = 3
        Else
            intNDec = 2
        End If
        LeggiNDecimali = intNDec
    End If
End Function

'Indica se la divisa passata è di Tipo EURO oppure no
Public Function MonetaEURO(ByVal CodDivisa&) As Boolean
    Dim q%, hSS As CRecordset
    Static UltCodDivisa&
    Static UltTipoEuro As Boolean
    
    If CodDivisa& = MXNU.CodCambioLire Or CodDivisa& = MXNU.CodCambioEuro Then
        MonetaEURO = True
        If CodDivisa& = MXNU.CodCambioEuro Then
            MCambioEuro = 1
        Else
            MCambioEuro = MXNU.CambioLireEuro
        End If
        UltCodDivisa = CodDivisa
        UltTipoEuro = MonetaEURO
    Else
        If UltCodDivisa <> CodDivisa Then
            Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT TipoEuro, CambioEuro FROM TabCambi WHERE Codice=" & CodDivisa)
            If Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
                MonetaEURO = (MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "TipoEuro", 0) = 1)
                MCambioEuro = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "CambioEuro", 1)
            Else
                MonetaEURO = False
                MCambioEuro = 1
            End If
            q = MXDB.dbChiudiSS(hSS)
            UltCodDivisa = CodDivisa
            UltTipoEuro = MonetaEURO
        Else
            MonetaEURO = UltTipoEuro
        End If
    End If
    
End Function

'------------------------------------------------------------
'nome:          LeggiCodiceCambio
'descrizione:
'parametri:     1. numero listino di cui leggere il cambio
'               2. flag considera listini trasformazione si/no
'ritorno:       codice cambio del listino passato
'annotazioni:
'------------------------------------------------------------
Public Function LeggiCodiceCambio(ByVal lngNrListino As Long, Optional ByVal bolListiniTrasf As Boolean = False) As Long
Dim strSQL As String
Dim hTr As MXKit.CRecordset
Dim lngCambio As Long
Dim strTabella As String

    On Local Error GoTo LeggiCodiceCambio_ERR
    strTabella = "TABLISTINI"
    If bolListiniTrasf Then
        strTabella = strTabella & "TRASFORMAZIONE"
    End If
    strSQL = "select CODCAMBIO from " & strTabella & " where NRLISTINO=" & lngNrListino
    Set hTr = MXDB.dbCreaSS(hndDBArchivi, strSQL)
    lngCambio = MXDB.dbGetCampo(hTr, TIPO_SNAPSHOT, "CODCAMBIO", 0)
    Call MXDB.dbChiudiSS(hTr)

LeggiCodiceCambio_END:
    LeggiCodiceCambio = lngCambio
    On Local Error GoTo 0
    Exit Function

LeggiCodiceCambio_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    lngCambio = 0
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("LeggiCodiceCambio", lngErrCod, strErrDsc))
    Resume LeggiCodiceCambio_END
End Function

