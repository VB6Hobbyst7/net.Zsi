Attribute VB_Name = "MPrtBlastBaf"
Option Explicit
DefInt A-Z

'/* orientation selections */
'#define DMORIENT_PORTRAIT   1
'#define DMORIENT_LANDSCAPE  2
Const DMORIENT_PORTRAIT = 1
Const DMORIENT_LANDSCAPE = 2

'parametri
'   stampa$:  nome del file con la lista delle stampe disponibili:
'             usare "Moduli" per i documenti
'   Nomi_Stampe()   vettore da caricare con i nomi dei file di stampa di
'                   Crystal Report
'   Nomi_SubReport Vettore contenente l'elenco degli eventuali SubReport all'interno del report n-esimo
'   ogg:  Controllo ComboBox nel quale caricare i nomi e le descrizioni dei file .rpt
'         contenuti in <stampa>.lst
'   sovrascr%:  forza la sovrascrittura del file .lst in ogni caso
'descrizione
'   carica Nomi_Stampe() e "ogg" con i nomi e le descrizioni dei file .rpt contenuti in <stampa>.lst;
'   se il file non esiste viene creato usando  il motore di Crystal Report
Sub CaricaListaStampe(ByVal Stampa As String, Ogg As Control, Nomi_Stampe() As String, Nomi_SubReport() As String, ByVal Sovrascr As Boolean)
    Call MXCREP.CaricaListaStampe(Stampa, Ogg, Nomi_Stampe(), Nomi_SubReport(), Sovrascr)
'    Dim nomestp$, q%
'    Dim job%, texth%, TextLen%, titolo$, Moduli%
'    Dim s$, hndf%, IndStp%
'    Dim filtrorpt$
'    Dim filelst$
'    Dim PathDirStampe As String
'    Dim colNomiStampe As New Collection, idxP As Integer
'    Dim strSR As String  'Elenco SubReport
''####    Dim strListaRepAltreDitte As String
'    Dim strListaDitte As String
'    Dim strDitta As Variant
'    Dim idxPathPers As Integer, idxPathPgm As Integer
'    Dim bolAggiungi As Boolean
'
'    If Trim(Stampa) = "" Then
'        Set colNomiStampe = Nothing
'        Exit Sub
'    End If
'
'    strListaDitte = MXNU.LeggiProfilo(MXNU.PercorsoLocal & "\DITTE.INI", "DITTE", vbNull, "")
'    ReDim percorsorpt(0 To UBound(Split(strListaDitte, Chr$(0))) + 2) As String
'    ReDim vetSegnaPostoPath(0 To UBound(Split(strListaDitte, Chr$(0))) + 2) As String
'
'    'ReDim percorsorpt(1 To 3) As String
'    'ReDim vetSegnaPostoPath(1 To 3) As String
'
'    Screen.MousePointer = vbHourglass
'    If LCase$(Right$(Stampa, 6)) = "moduli" Then
'        Moduli = 0
'    ElseIf LCase$(Right$(Stampa, 8)) = "ricevute" Then
'        Moduli = 1
'    ElseIf LCase$(Right$(Stampa, 7)) = "deleghe" Then
'        Moduli = 2
'    ElseIf LCase$(Right$(Stampa, 9)) = "scontrini" Then
'        Moduli = 3
'    ElseIf LCase$(Right$(Stampa, 6)) = "etcdoc" Then
'        Moduli = 5
'    ElseIf LCase$(Right$(Stampa, 7)) = "dichiva" Then
'        Moduli = 6
'    ElseIf LCase$(Right$(Stampa, 7)) = "ordprod" Then
'        Moduli = 7
'    Else
'        Moduli = 4
'    End If
'    For idxP = 0 To UBound(Split(strListaDitte, Chr$(0)))
'        strDitta = Split(strListaDitte, Chr$(0))(idxP)
'        vetSegnaPostoPath(idxP) = "%PATHPERSDITTA-" & strDitta & "%"
'        Select Case Moduli
'            Case 0    'Moduli Documenti
'                percorsorpt(idxP) = MXNU.PercorsoPers & "\" & strDitta & "\stampe\moduli\"
'            Case 1    'Ricevute
'                percorsorpt(idxP) = MXNU.PercorsoPers & "\" & strDitta & "\stampe\moduli\"
'            Case 2    'Deleghe Iva
'                percorsorpt(idxP) = MXNU.PercorsoPers & "\" & strDitta & "\stampe\moduli\"
'            Case 3    'Scontrini
'                percorsorpt(idxP) = MXNU.PercorsoPers & "\" & strDitta & "\stampe\moduli\"
'            Case 5    'Etic. Doc.
'                percorsorpt(idxP) = MXNU.PercorsoPers & "\" & strDitta & "\stampe\moduli\"
'            Case 6    'Dichiarazione Iva
'                percorsorpt(idxP) = MXNU.PercorsoPers & "\" & strDitta & "\stampe\moduli\"
'            Case 7     'Stampe ordini produzione
'                percorsorpt(idxP) = MXNU.PercorsoPers & "\" & strDitta & "\stampe\moduli\"
'            Case Else   'Laser
'                percorsorpt(idxP) = MXNU.PercorsoPers & "\" & strDitta & "\stampe\laser\"
'        End Select
'    Next idxP
'    idxPathPers = idxP
'    idxPathPgm = idxP + 1
'    vetSegnaPostoPath(idxPathPers) = "%PATHPERS%"
'    vetSegnaPostoPath(idxPathPgm) = "%PATHPGM%"
'    Select Case Moduli
'        Case 0    'Moduli Documenti
'            filelst = MXNU.PercorsoPgm & "\stampe\moduli\moduli.lst"
'            percorsorpt(idxPathPers) = MXNU.PercorsoPers & "\stampe\moduli\"
'            percorsorpt(idxPathPgm) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
'            PathDirStampe = "\moduli\"
'            filtrorpt = "*.rpt"
'        Case 1    'Ricevute
'            filelst = MXNU.PercorsoPgm & "\stampe\moduli\ricevute.lst"
'            percorsorpt(idxPathPers) = MXNU.PercorsoPers & "\stampe\moduli\"
'            percorsorpt(idxPathPgm) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
'            PathDirStampe = "\moduli\"
'            filtrorpt = "ricev*.rpt"
'        Case 2    'Deleghe Iva
'            filelst = MXNU.PercorsoPgm & "\stampe\moduli\deleghe.lst"
'            percorsorpt(idxPathPers) = MXNU.PercorsoPers & "\stampe\moduli\"
'            percorsorpt(idxPathPgm) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
'            PathDirStampe = "\moduli\"
'            filtrorpt = "deleg*.rpt"
'        Case 3    'Scontrini
'            filelst = MXNU.PercorsoPgm & "\stampe\moduli\scontr.lst"
'            percorsorpt(idxPathPers) = MXNU.PercorsoPers & "\stampe\moduli\"
'            percorsorpt(idxPathPgm) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
'            PathDirStampe = "\moduli\"
'            filtrorpt = "scontr*.rpt"
'        Case 5    'Etic. Doc.
'            filelst = MXNU.PercorsoPgm & "\stampe\moduli\etcdoc.lst"
'            percorsorpt(idxPathPers) = MXNU.PercorsoPers & "\stampe\moduli\"
'            percorsorpt(idxPathPgm) = MXNU.PercorsoStampe & "\moduli\"   'MXNU.PercorsoPgm
'            PathDirStampe = "\moduli\"
'            filtrorpt = "etcdoc*.rpt"
'        Case 6    'Dichiarazione Iva
'            filelst = MXNU.PercorsoPgm & "\stampe\moduli\dichiva.lst"
'            percorsorpt(idxPathPers) = MXNU.PercorsoPers & "\stampe\moduli\"
'            percorsorpt(idxPathPgm) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
'            PathDirStampe = "\moduli\"
'            filtrorpt = "dichiva*.rpt"
'        Case 7     'Stampe ordini produzione
'            filelst = MXNU.PercorsoPgm & "\stampe\moduli\ordprod.lst"
'            percorsorpt(idxPathPers) = MXNU.PercorsoPers & "\stampe\moduli\"
'            percorsorpt(idxPathPgm) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
'            PathDirStampe = "\moduli\"
'            filtrorpt = "ordprod*.rpt"
'        Case Else   'Laser
'            filelst = MXNU.PercorsoPgm & "\stampe\" & Stampa & ".lst"
'            percorsorpt(idxPathPers) = MXNU.PercorsoPers & "\stampe\laser\"
'            percorsorpt(idxPathPgm) = MXNU.PercorsoStampe & "\laser\"     'MXNU.PercorsoPgm
'            PathDirStampe = "\laser\"
'            filtrorpt = Stampa & "*.rpt"
'    End Select
'
''********* VECCHIA GESTIONE PRE-SVILUPPO 999 *************************************************************************
''    vetSegnaPostoPath(1) = "%PATHPERSDITTA-" & MXNU.DittaAttiva & "%"
''    vetSegnaPostoPath(2) = "%PATHPERS%"
''    vetSegnaPostoPath(3) = "%PATHPGM%"
''    Select Case Moduli
''        Case 0    'Moduli Documenti
''            filelst = MXNU.PercorsoPgm & "\stampe\moduli\moduli.lst"
''            percorsorpt(1) = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\stampe\moduli\"
''            percorsorpt(2) = MXNU.PercorsoPers & "\stampe\moduli\"
''            percorsorpt(3) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
''            PathDirStampe = "\moduli\"
''            filtrorpt = "*.rpt"
''        Case 1    'Ricevute
''            filelst = MXNU.PercorsoPgm & "\stampe\moduli\ricevute.lst"
''            percorsorpt(1) = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\stampe\moduli\"
''            percorsorpt(2) = MXNU.PercorsoPers & "\stampe\moduli\"
''            percorsorpt(3) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
''            PathDirStampe = "\moduli\"
''            filtrorpt = "ricev*.rpt"
''        Case 2    'Deleghe Iva
''            filelst = MXNU.PercorsoPgm & "\stampe\moduli\deleghe.lst"
''            percorsorpt(1) = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\stampe\moduli\"
''            percorsorpt(2) = MXNU.PercorsoPers & "\stampe\moduli\"
''            percorsorpt(3) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
''            PathDirStampe = "\moduli\"
''            filtrorpt = "deleg*.rpt"
''        Case 3    'Scontrini
''            filelst = MXNU.PercorsoPgm & "\stampe\moduli\scontr.lst"
''            percorsorpt(1) = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\stampe\moduli\"
''            percorsorpt(2) = MXNU.PercorsoPers & "\stampe\moduli\"
''            percorsorpt(3) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
''            PathDirStampe = "\moduli\"
''            filtrorpt = "scontr*.rpt"
''        Case 5    'Etic. Doc.
''            filelst = MXNU.PercorsoPgm & "\stampe\moduli\etcdoc.lst"
''            percorsorpt(1) = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\stampe\moduli\"
''            percorsorpt(2) = MXNU.PercorsoPers & "\stampe\moduli\"
''            percorsorpt(3) = MXNU.PercorsoStampe & "\moduli\"   'MXNU.PercorsoPgm
''            PathDirStampe = "\moduli\"
''            filtrorpt = "etcdoc*.rpt"
''        Case 6    'Dichiarazione Iva
''            filelst = MXNU.PercorsoPgm & "\stampe\moduli\dichiva.lst"
''            percorsorpt(1) = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\stampe\moduli\"
''            percorsorpt(2) = MXNU.PercorsoPers & "\stampe\moduli\"
''            percorsorpt(3) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
''            PathDirStampe = "\moduli\"
''            filtrorpt = "dichiva*.rpt"
''        Case 7     'Stampe ordini produzione
''            filelst = MXNU.PercorsoPgm & "\stampe\moduli\ordprod.lst"
''            percorsorpt(1) = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\stampe\moduli\"
''            percorsorpt(2) = MXNU.PercorsoPers & "\stampe\moduli\"
''            percorsorpt(3) = MXNU.PercorsoStampe & "\moduli\"    'MXNU.PercorsoPgm
''            PathDirStampe = "\moduli\"
''            filtrorpt = "ordprod*.rpt"
''        Case Else   'Laser
''            filelst = MXNU.PercorsoPgm & "\stampe\" & Stampa & ".lst"
''            percorsorpt(1) = MXNU.PercorsoPers & "\" & MXNU.DittaAttiva & "\stampe\laser\"
''            percorsorpt(2) = MXNU.PercorsoPers & "\stampe\laser\"
''            percorsorpt(3) = MXNU.PercorsoStampe & "\laser\"     'MXNU.PercorsoPgm
''            PathDirStampe = "\laser\"
''            filtrorpt = Stampa & "*.rpt"
''    End Select
''*********************************************************************************************************************
'
'    IndStp = 0  'Contatore Stampe per Array dei Nomi_Stampe
'    Ogg.Clear
'    If Dir(filelst, vbNormal) <> "" And Sovrascr Then
'        'Memorizzo le righe del file relative a report di altre ditte.
'        'Rif. Sviluppo Nr. 999
''#####        strListaRepAltreDitte = LeggiReportAltreDitte(filelst)
'        Kill filelst
'    End If
'    If Dir(filelst, vbNormal) = "" Then
'        q = vbYes
'        If Not Sovrascr Then
'            Call MXNU.MsgBoxEX(1118, vbOKOnly + vbExclamation, 1007)
'        End If
'        If q = vbYes Then
'            Dim objCRW As MXKit.CCrw
'            Dim strRepDitta As Variant
'            Set objCRW = MXCREP.CreaCCrw()
'
'            hndf = FreeFile
'            Open filelst$ For Output Shared As hndf
'            On Local Error Resume Next
'            'For idxP = 1 To 3
'            For idxP = 0 To UBound(percorsorpt)
'                nomestp = Dir(percorsorpt(idxP) & filtrorpt, vbNormal)
'                Do While nomestp <> ""
'                    If (Moduli <> 0) Or (Moduli = 0 And InStr("ricev,deleg", LCase(Left(nomestp, 5))) = 0 And InStr("scontr,etcdoc", LCase(Left(nomestp, 6))) = 0 And InStr("dichiva,ordprod", LCase(Left(nomestp, 7))) = 0) Then
'                        If objCRW.LeggiReportInfo(percorsorpt(idxP) & nomestp, titolo, 0, 0) Then
'                            If Moduli = 0 Or Moduli = 1 Or Moduli = 5 Then
'                                s = LTrim(titolo)
'                            Else
'                                If IsNumeric(Right(Replace(UCase(nomestp), ".RPT", ""), 2)) Then
'                                    s = Right(Replace(UCase(nomestp), ".RPT", ""), 2) & "  " & titolo
'                                Else
'                                    s = LTrim(titolo)
'                                End If
'                            End If
'                            MXNU.MostraMsgInfo 70000, LTrim(titolo)
'                            'colNomiStampe.Add nomestp, UCase(nomestp)
'                            'If Err = 0 Then
'                                If (idxP = idxPathPers) Or (idxP = idxPathPgm) Or (InStr(percorsorpt(idxP), MXNU.DittaAttiva) <> 0) Then
'                                    bolAggiungi = Not EsisteElementoCollection(colNomiStampe, UCase(nomestp))
'                                    If bolAggiungi Then
'                                        Ogg.addItem s
'                                        Ogg.ItemData(Ogg.NewIndex) = IndStp
'                                        'If idxP = 3 Then   'PathPgm -> Comprende già la parola "stampe"
'                                        If idxP = idxPathPgm Then   'PathPgm -> Comprende già la parola "stampe"
'                                            Nomi_Stampe(IndStp) = vetSegnaPostoPath(idxP) & PathDirStampe & nomestp
'                                        Else
'                                            Nomi_Stampe(IndStp) = vetSegnaPostoPath(idxP) & "\stampe" & PathDirStampe & nomestp
'                                        End If
'                                        colNomiStampe.Add nomestp, UCase(nomestp)
'                                    End If
'                                End If
'                                'Lettura eventuali SubReport
'                                strSR = objCRW.ListaSubReport(percorsorpt(idxP) & "\" & nomestp)
'                                If (idxP = idxPathPers) Or (idxP = idxPathPgm) Or (InStr(percorsorpt(idxP), MXNU.DittaAttiva) <> 0) Then
'                                    If bolAggiungi Then Nomi_SubReport(IndStp) = strSR
'                                End If
'
'                                'If idxP = 3 Then
'                                If idxP = idxPathPgm Then
'                                    Print #hndf, vetSegnaPostoPath(idxP) & PathDirStampe & nomestp & "-" & s & ";" & strSR
'                                Else
'                                    Print #hndf, vetSegnaPostoPath(idxP) & "\stampe" & PathDirStampe & nomestp & "-" & s & ";" & strSR
'                                End If
'                                If (idxP = idxPathPers) Or (idxP = idxPathPgm) Or (InStr(percorsorpt(idxP), MXNU.DittaAttiva) <> 0) Then
'                                    If bolAggiungi Then IndStp = IndStp + 1
'                                End If
'                            'End If
'                        End If
'                    End If
'                    q = DoEvents()
'                    nomestp = Dir
'                Loop
''#####################################################################################################################
''                If strListaRepAltreDitte <> "" Then
''                    'Reinserisco nel file lst la lista dei report personalizzati di altre ditte (Sviluppo 999)
''                    For Each strRepDitta In Split(strListaRepAltreDitte, vbCrLf)
''                        Print #hndf, strRepDitta
''                    Next strRepDitta
''                End If
''#####################################################################################################################
'            Next idxP
'            Close hndf
'            Set objCRW = Nothing
'        End If
'    Else
'        'caricamento lista dei file di stampa da Stampa$ & ".LST"
'        Dim bolValido As Boolean
'        Dim NomeRpt As String
'        hndf = FreeFile
'        Open filelst For Input Shared As hndf
'        Do While Not EOF(hndf)
'            Line Input #hndf, nomestp
'            If Trim(nomestp) <> "" Then  'Rif. 897
'                If InStr(nomestp, ";") Then
'                    strSR = Mid(nomestp, InStr(nomestp, ";") + 1)
'                    nomestp = Left(nomestp, InStr(nomestp, ";") - 1)
'                Else
'                    strSR = ""
'                End If
'                If InStr(nomestp, "%PATHPERSDITTA%") Then   'Altrimenti i file lst con vecchia gestione non visualizzano i report se presenti per la ditta attiva
'                    bolValido = CercaRptDittaAtt(Trim(Left(nomestp, InStr(nomestp, "-") - 1)))
'                ElseIf InStr(nomestp, "%PATHPERSDITTA-" & MXNU.DittaAttiva & "%") <> 0 Then  'Report Personalizzato per la ditta attiva: controllo se esiste il file
'                    bolValido = CercaRptDittaAtt(Trim(Left(nomestp, InStrRev(nomestp, "-") - 1)))
'                ElseIf InStr(nomestp, "%PATHPERSDITTA") <> 0 Then 'Report personalizzato di altra ditta: lo escludo dalla lista
'                    bolValido = False
'                Else
'                    bolValido = True
'                End If
'                If bolValido Then
'                    NomeRpt = Mid(nomestp, InStrRev(nomestp, "\") + 1)
'                    NomeRpt = Left(NomeRpt, InStr(LCase(NomeRpt), ".rpt") + 3)
'                    If Not EsisteElementoCollection(colNomiStampe, UCase(NomeRpt)) Then
'                        If InStr(nomestp, "%PATHPERSDITTA-") <> 0 Then
'                            q = InStrRev(nomestp, "-")
'                        Else
'                            q = InStr(nomestp, "-")
'                        End If
'                        Ogg.addItem Mid$(nomestp, q + 1)
'                        Ogg.ItemData(Ogg.NewIndex) = IndStp
'                        Nomi_Stampe(IndStp) = Left$(nomestp, q - 1)
'                        Nomi_SubReport(IndStp) = strSR
'                        IndStp = IndStp + 1
'                        colNomiStampe.Add NomeRpt, UCase(NomeRpt)
'                    End If
'                End If
'            End If
'        Loop
'        Close hndf
'    End If
'
'    MXNU.MostraMsgInfo ""
'    Screen.MousePointer = vbDefault
'    Set colNomiStampe = Nothing
End Sub

'Private Function CercaRptDittaAtt(ByVal strNomeStampa As String) As Boolean
'    On Local Error Resume Next
'    Dim strNomeFile$
'    If InStr(strNomeStampa, "%PATHPERSDITTA%") <> 0 Then
'        strNomeFile = Replace(strNomeStampa, "%PATHPERSDITTA%", MXNU.PercorsoPers & "\" & MXNU.DittaAttiva)
'    Else
'        strNomeFile = Replace(strNomeStampa, "%PATHPERSDITTA-" & MXNU.DittaAttiva & "%", MXNU.PercorsoPers & "\" & MXNU.DittaAttiva)
'    End If
'    CercaRptDittaAtt = (Dir(strNomeFile, vbNormal) <> "")
'    On Local Error GoTo 0
'End Function


'Private Function LeggiReportAltreDitte(ByVal NomeFileLst As String) As String
'    Dim objFSO As New Scripting.FileSystemObject
'    Dim txtStream As TextStream
'    Dim strTxtLine As String
'    Dim strReportAltreDitte As String
'    On Local Error GoTo err_LeggiReportAltreDitte
'    Set txtStream = objFSO.OpenTextFile(NomeFileLst, ForReading, False)
'    strReportAltreDitte = ""
'    Do While Not txtStream.AtEndOfStream
'        strTxtLine = txtStream.ReadLine
'        If InStr(strTxtLine, "%PATHPERSDITTA-") <> 0 Then
'            If InStr(strTxtLine, "%PATHPERSDITTA-" & MXNU.DittaAttiva & "%") = 0 Then
'                strReportAltreDitte = strReportAltreDitte & strTxtLine & vbCrLf
'            End If
'        End If
'    Loop
'    If strReportAltreDitte <> "" Then strReportAltreDitte = Left(strReportAltreDitte, Len(strReportAltreDitte) - 2)
'
'Esci_LeggiReportAltreDitte:
'    LeggiReportAltreDitte = strReportAltreDitte
'    txtStream.Close
'    Set objFSO = Nothing
'    On Local Error GoTo 0
'    Exit Function
'
'err_LeggiReportAltreDitte:
'    MXNU.MsgBoxEX 1009, vbExclamation, 1007, Array("LeggiReportAltreDitte", Err.Number, Err.Description)
'    Resume Esci_LeggiReportAltreDitte
'
'End Function
'
'Posiziona il combo con la lista stampe sulla prima stampa che contiene la dicitura euro
'Rif. Sk. Sviluppi Nr. 752
Public Sub PosizioneStpEuro(oggCmb As ComboBox)
    '>>>> SPOSTATO FUNZIONE SU AMBCRW DEL KIT PER POTERLA USARE ANCHE NELLE ESTENSIONI
    Call MXCREP.PosizioneStpEuro(oggCmb)
'    Dim i As Integer
'    Dim strElem As String
'    Dim bolTrovato As Boolean
'    If oggCmb.ListCount = 0 Then Exit Sub
'    For i = 0 To oggCmb.ListCount - 1
'        strElem = oggCmb.List(i)
'        If InStrRev(UCase(strElem), " EURO ") Or InStrRev(UCase(strElem), "(EURO)") Then
'            bolTrovato = True
'        ElseIf InStrRev(UCase(strElem), " EURO") Then
'            bolTrovato = (Mid(strElem, InStrRev(UCase(strElem), " EURO") + 5) = ")" Or Mid(strElem, InStrRev(UCase(strElem), " EURO") + 5) = "" Or Mid(strElem, InStrRev(UCase(strElem), " EURO") + 5) = "(" Or Mid(strElem, InStrRev(UCase(strElem), " EURO") + 5) = ".")
'        End If
'        If bolTrovato Then Exit For
'    Next i
'    If bolTrovato Then
'        oggCmb.ListIndex = i
'    Else
'        oggCmb.ListIndex = 0
'    End If
End Sub

