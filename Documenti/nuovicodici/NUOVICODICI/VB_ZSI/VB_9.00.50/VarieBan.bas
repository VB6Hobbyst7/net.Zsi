Attribute VB_Name = "VarieSio"
Option Explicit
DefLng A-Z

'===============================================
'       definizione costanti globali
'===============================================
'indici schede particolari
Global Const ID_SCHEDA_FORM = -1
Global Const ID_SCHEDA_SITUAZIONE = -2
'costanti accessi
Global Const ACC_NONDEFINITO = -1
Global Const ACC_NESSUNO = 0
Global Const ACC_LETTURA = 1
Global Const ACC_MODIFICA = 2
Global Const ACC_INSERISCI = 4
Global Const ACC_ANNULLA = 8
Global Const ACC_TUTTI = 15

Const MNU_ITEM_GESTIONE_ACCESSI = 3

'===============================================
'       definizione variabili
'===============================================
Dim mStrBufferModuliChiave As String
Dim mStrLogFile As String
Dim cLogFile As String


'===============================================
'       definizione variabili specifiche
'===============================================
Global RowSpreadComp As Long
Global strPathImmagine As String


Public Function AccessoLettura(ByVal intAccesso) As Boolean
    AccessoLettura = ((intAccesso And ACC_LETTURA) = ACC_LETTURA)
End Function

Public Function AccessoModifica(ByVal intAccesso) As Boolean
    AccessoModifica = ((intAccesso And ACC_MODIFICA) = ACC_MODIFICA)
End Function

Public Function AccessoInserimento(ByVal intAccesso) As Boolean
    AccessoInserimento = ((intAccesso And ACC_INSERISCI) = ACC_INSERISCI)
End Function

Public Function AccessoAnnulla(ByVal intAccesso) As Boolean
    AccessoAnnulla = ((intAccesso And ACC_ANNULLA) = ACC_ANNULLA)
End Function

'NOME           : FormImpostaAccessi
'DESCRIZIONE    : legge ed imposta gli accessi per la form
'PARAMETRO 1    : form di cui definire gli accessi
'PARAMETRO 2    : maschera per i bottoni della toolbar (viene modificata in base agli accessi)
'RITORNO        : maschera accessi per la form (gli accessi possono essere letti in seguito mediante le funzioni
'                 AccessiLettura,AccessiModifica,AccessiInserimento,AccessiAnnullamento)
Public Function FormImpostaAccessi(ByVal frmDef As Form, lngButtonMask As Long) As Integer
Dim lngFormID As Long
Dim ctrGen As Control, ctrGen1 As Control
Dim intAccesso As Integer
Dim intAccLing As Integer
    
    If (Not (frmDef Is Nothing) And MXNU.CtrlAccessi) Then
        On Local Error Resume Next
        lngFormID = frmDef.HelpContextID
        'leggo l'accesso per la form...
        intAccesso = LeggiAccessi(MXNU.UtenteAttivo, lngFormID, ID_SCHEDA_FORM)
        If (intAccesso = ACC_NESSUNO) Then
            lngButtonMask = 0
        Else
            If (intAccesso And ACC_INSERISCI) = 0 Then lngButtonMask = lngButtonMask And (Not BTN_INS_MASK)
            If (intAccesso And ACC_MODIFICA) = 0 Then lngButtonMask = lngButtonMask And (Not BTN_REG_MASK)
            If (intAccesso And ACC_ANNULLA) = 0 Then lngButtonMask = lngButtonMask And (Not BTN_ANN_MASK)
        End If
        '...e per le linguette
        For Each ctrGen In frmDef
            If (TypeName(ctrGen) = "MWLinguetta" And ctrGen.Name = "Ling") Then
                intAccLing = LeggiAccessi(MXNU.UtenteAttivo, lngFormID, ctrGen.Index)
                If (intAccLing = ACC_NESSUNO) Then
                    'nessun accesso -> rendo invisibile la linguetta e la scheda
                    frmDef.Ling(ctrGen.Index).Enabled = False
                    frmDef.Scheda(ctrGen.Index).Enabled = False
                    For Each ctrGen1 In frmDef.Scheda(ctrGen.Index).Controls
                        ctrGen1.Visible = False
                    Next ctrGen1
                ElseIf (intAccLing = ACC_LETTURA) Then
                    'accesso sola lettura -> abilito la linguetta e disabilito la scheda
                    frmDef.Ling(ctrGen.Index).Enabled = True
                    'frmDef.Scheda(ctrGen.Index).Enabled = False
                    For Each ctrGen1 In frmDef.Scheda(ctrGen.Index).Controls
                        ctrGen1.Enabled = False
                    Next ctrGen1
                    frmDef.Scheda(ctrGen.Index).TabStop = True
                    frmDef.Scheda(ctrGen.Index).TabIndex = frmDef.Ling(ctrGen.Index).TabIndex + 1
                End If
            End If
        Next
        On Local Error GoTo 0
    Else
        intAccesso = ACC_TUTTI
    End If
    FormImpostaAccessi = intAccesso
End Function

'Utente appartenente ad un gruppo
'GRUPPO     D       D       N       N
'UTENTE     D       N       N       D
'ACCESSI   G&U      G      ALL    ALL&U
'           1a      2a      3a      4a
'Utente NON appartenente a gruppi
'UTENTE     D       N
'ACCESSI    U       NO
'           1b      2b
Function LeggiAccessi(ByVal strUtente As String, ByVal lngFormID As Long, ByVal intScheda As Integer) As Integer
Dim intq As Integer
Dim strSQL As String
Dim hSS As CRecordSet
Dim bolEnd As Boolean
Dim intAccGrp As Integer
Dim intAccUsr As Integer

    intAccUsr = ACC_NONDEFINITO
    'leggo gli accessi per l'utente
    strSQL = "SELECT TipoAccesso" _
            & " FROM TabAccessiUtente" _
            & " WHERE CodUtente = '" & strUtente & "'" _
            & " AND HelpID = " & lngFormID _
            & " AND IndiceScheda = " & intScheda
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    If (Not bolEnd) Then intAccUsr = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "TipoAccesso", ACC_NESSUNO)
    intq = MXDB.dbChiudiSS(hSS)
    'leggo gli accessi per i gruppi di cui l'utente fa parte
    strSQL = "SELECT CodGruppo" _
            & " FROM TabMembriGruppo" _
            & " WHERE  CodUtente = '" & strUtente & "'"
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
    intAccGrp = ACC_NONDEFINITO
    If (MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
        If (intAccUsr = ACC_NONDEFINITO) Then
            LeggiAccessi = ACC_NESSUNO '(2b)
        Else
            LeggiAccessi = intAccUsr '(1b)
        End If
    Else
        intAccGrp = ACC_NONDEFINITO
        intq = MXDB.dbChiudiSS(hSS)
        strSQL = "SELECT AG.TipoAccesso" _
                    & " FROM {oj TabAccessiGruppo AG INNER JOIN TabMembriGruppo MG" _
                    & " ON AG.CodGruppo=MG.CodGruppo" _
                    & " AND MG.CodUtente = " & hndDBArchivi.FormatoSQL(strUtente, DB_TEXT) _
                    & " AND AG.HelpID = " & lngFormID _
                    & " AND AG.IndiceScheda = " & intScheda & "}"
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
        bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
        If (Not bolEnd) Then
            intAccGrp = ACC_NESSUNO
            Do
                intAccGrp = intAccGrp Or MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "TipoAccesso", ACC_TUTTI)
                bolEnd = Not MXDB.dbSuccessivo(hSS, TIPO_SNAPSHOT)
            Loop While (Not bolEnd)
        End If
        If (intAccGrp = ACC_NONDEFINITO) Then
            If (intAccUsr = ACC_NONDEFINITO) Then
                LeggiAccessi = ACC_TUTTI '(3a)
            Else
                LeggiAccessi = (ACC_TUTTI And intAccUsr) '(4a)
            End If
        Else
            If (intAccUsr = ACC_NONDEFINITO) Then
                LeggiAccessi = intAccGrp '(2a)
            Else
                LeggiAccessi = (intAccGrp And intAccUsr) '(1a)
            End If
        End If
    End If
    intq = MXDB.dbChiudiSS(hSS)
    
End Function


Function ChkLoop() As Boolean
    Dim cSql As String
    Dim rSql As CRecordSet
    Dim cCodArt As String
    Dim oFSO As New FileSystemObject
    
    ChkLoop = False
    
    cSql = "SELECT ID, CODARTF FROM ECOMTREE2 WHERE IDP = IDF AND TIPOF <> 'R'"
    
    Set rSql = MXDB.dbCreaSS(hndDBArchivi, cSql)

    cCodArt = ""

    Do While Not MXDB.dbFineTab(rSql)

        cCodArt = cCodArt & vbCrLf & MXDB.dbGetCampo(rSql, TIPO_SNAPSHOT, "CODARTF", "")

        DoEvents
        
        Call MXDB.dbSuccessivo(rSql)
        
    Loop

    Call MXDB.dbChiudiSS(rSql)
    
    If cCodArt <> "" Then
        
        MXNU.MsgBoxEX "Rilevati i seguenti nodi in loop: " & vbCrLf & cCodArt & vbCrLf & vbCrLf & "Gli articoli elencati saranno eliminati e sarà necessario reinserirli", vbExclamation, "Controllo Nodi in Loop"
        
        cLogFile = MXNU.PercorsoPreferenze & "\Loop.log"
        
        If oFSO.FileExists(cLogFile) Then
            
            Call MXNU.ImpostaErroriSuLog(cLogFile, False)
        
        Else
        
            Call MXNU.ImpostaErroriSuLog(cLogFile, True)
            
        End If
        
        MXNU.MsgBoxEX Now() & vbCrLf & vbCrLf & "Rilevati i seguenti nodi in loop: " & vbCrLf & cCodArt & vbCrLf & vbCrLf & "Gli articoli elencati saranno eliminati e sarà necessario reinserirli", vbExclamation, "Controllo Nodi in Loop"
        
        Call MXNU.ChiudiErroriSuLog
        
        cSql = "EXEC dbo.ECOM_DEL_LOOP"
        
        MXDB.dbEseguiSQL hndDBArchivi, cSql
        
        ChkLoop = True
    
    End If
    
End Function




