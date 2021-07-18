Attribute VB_Name = "MVarie"
Option Explicit
DefInt A-Z

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Dati imponibile e aliquota (in percentuale) ritorna l'imposta arrotondata per eccesso
Function IVANormale(ByVal Impon As Variant, ByVal alq As Variant, ByVal cv_dec As Long, ByVal CodCambio As Long) As Variant

    Dim Imposta As Variant

    If CodCambio = MXNU.CodCambioLire Then
        'Imposta = Abs(Impon) / 100 * alq * cv_dec
        Imposta = Abs(Impon) / 100 * alq
        If Fix(Imposta) <> Imposta Then
            If cv_dec = 0 Then
                Imposta = Fix(Imposta) + 1
            Else
                Imposta = fdec(Imposta, cv_dec)
            End If
        End If
        'IVANormale = CDec((Sgn(Impon) * Imposta / cv_dec))
        IVANormale = CDec((Sgn(Impon) * Imposta))
    Else
        IVANormale = fdec(Impon / 100 * alq, cv_dec)
    End If
End Function


Function WindowsDirectory() As String
    Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, OSGetWindowsDirectory(WinPath, Len(WinPath)))
End Function

Public Sub FrmZummaControlli(frm As Form, ByVal Zum As Variant)

    Dim ctrl As Control
    Dim autosz As Variant

    On Local Error Resume Next

    frm.Height = CInt(frm.Height * Zum)
    frm.width = CInt(frm.width * Zum)
    For Each ctrl In ControlliForm(frm)
        If TypeName(ctrl) <> "fpSpread" Then
            ctrl.Col = -1
            ctrl.Row = -1
            autosz = ctrl.AutoSize
            ctrl.AutoSize = False
            ctrl.ReDraw = False
        End If
        ctrl.FontSize = CInt(ctrl.FontSize * Zum)
        ctrl.Top = CInt(ctrl.Top * Zum)
        ctrl.Left = CInt(ctrl.Left * Zum)
        ctrl.Height = CInt(ctrl.Height * Zum)
        ctrl.width = CInt(ctrl.width * Zum)
        If TypeName(ctrl) <> "fpSpread" Then
            ctrl.AutoSize = autosz
            ctrl.ReDraw = True
        End If

    Next


End Sub

Sub CentraInContenitore(objOggetto As Object, ByVal enmOrientation As MSComCtl2.OrientationConstants)
    If enmOrientation = cc2OrientationHorizontal Then
        objOggetto.Left = (objOggetto.Container.ScaleWidth - objOggetto.width) \ 2
    Else
        objOggetto.Top = (objOggetto.Container.ScaleHeight - objOggetto.Height) \ 2
    End If
End Sub

Sub OrdinaArray(IndArray() As Long, KeyArray() As Variant)
    Dim span&, num&, i&, j&
    Dim Supp As Variant
    Dim IndSupp As Long
    Dim q%
    num = UBound(KeyArray)
    span& = num& \ 2
    Do While span& > 0
        For i& = span& To num& - 1
            For j& = (i& - span& + 1) To 1 Step -span&
                If KeyArray(j&) <= KeyArray(j& + span&) Then Exit For
                Supp = KeyArray(j&)
                KeyArray(j&) = KeyArray(j& + span&)
                KeyArray(j& + span&) = Supp
                        
                IndSupp& = IndArray(j&)
                IndArray(j&) = IndArray(j& + span&)
                IndArray(j& + span&) = IndSupp&
            Next j&
        Next i&
        span& = span& \ 2
    Loop
End Sub

Sub CaricaFileSuCombo(Cmb As ComboBox, strPath As String, strWildCard As String)
    'riempie il combo dei files contenuti nella directory indicata, controllando le wildcard.
    Dim objFSO As Scripting.FileSystemObject
    Dim objDir As Folder
    Dim objFile As file
    Dim objFiles As Files

    On Local Error GoTo CaricaFileSuCombo_ERR
    Set objFSO = New Scripting.FileSystemObject
    Cmb.Clear
    Cmb.addItem " "
    If objFSO.FolderExists(strPath) Then
        Set objDir = objFSO.GetFolder(strPath)
        Set objFiles = objDir.Files
        For Each objFile In objFiles
            If UCase(Right(objFile.NAME, Len(strWildCard))) = strWildCard Then
                Cmb.addItem objFile.NAME
            End If
        Next
        Set objDir = Nothing
        Set objFile = Nothing
        Set objFiles = Nothing
    Else
        'Manca la directory
        Call MXNU.MsgBoxEX(MXNU.CaricaStringaRes(2231), vbOKOnly + vbExclamation, MXNU.CaricaStringaRes(2092))
    End If

FineSub:
    On Local Error GoTo 0
    Set objFSO = Nothing
    Set objDir = Nothing
    Set objFiles = Nothing
    Set objFile = Nothing
    Exit Sub
CaricaFileSuCombo_ERR:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Local Error Resume Next
    'Call MXDB.dbRollBack(hndDBArchivi)
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("CaricaFileSuCombo", errNumber, errDescription))
    Resume Next
End Sub

'Lancia una chiamata al numero specificato attraverso VOIspeed. Se mostraMessaggi=true allora
'la funzione, in caso di errori, visualizza anche i messaggi a video. Restituisce true se
'la chiamata ha successo, false altrimenti
Function ChiamaTelefonoVoispeed(ByVal numero As String, ByVal mostraMessaggi As Boolean) As Boolean
    Dim Ret As Long
    
    'Verifica che il numero non sia vuoto
    numero = Replace(numero, " ", "")
    If numero = "" Then
        ChiamaTelefonoVoispeed = False
        Exit Function
    End If
    
    'Verifica il formato del numero
    If Len(numero) > 20 Then
        If mostraMessaggi Then
            Call MXNU.MsgBoxEX(3272, vbExclamation, 2092)
        End If
    
        ChiamaTelefonoVoispeed = False
        Exit Function
    End If
    
    'Normalizza il numero
    numero = Replace(numero, """", "")
    numero = """" & numero & """"
    
    'Avvia il programma voispeed attraverso un link apposito
    Ret = ShellExecute(0, "Open", "voispeed:" & numero, "", "", 0)
    
    If Ret > 32 Then
        ChiamaTelefonoVoispeed = True
        Exit Function
    Else
        If mostraMessaggi Then
            Call MXNU.MsgBoxEX(3273, vbInformation, 2092)
        End If
        
        ChiamaTelefonoVoispeed = False
        Exit Function
    End If
End Function

'Restituisce true se su mw.ini è stato attivato voispeed (con voispeed=1), false altrimenti
Function IsVoispeedActive() As Boolean
    Dim i As Integer
    
    'Prova a leggere dall'impostazione utente oppure globale
    i = Cast(MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", MXNU.UtenteSistema, "voispeed", "-1"), vbInteger)
    If i = -1 Then
        i = Cast(MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "METODOW", "voispeed", "-1"), vbInteger)
    End If
    
    'Se non è stato impostato lo mette per default a disattivato
    If i > 0 Then
        IsVoispeedActive = True
    Else
        IsVoispeedActive = False
    End If
End Function

Function GetCtConfigurazione(nomeConfigurazione As String, valoreDefault As String) As String
Dim strValore As String
Dim strQuery As String

Dim HSS As MXKit.CRecordSet

    On Local Error GoTo ERR_GetCtConfigurazione
    
    strValore = valoreDefault

    strQuery = "SELECT Valore FROM zzCtConfigurazioni WHERE Nome = " + hndDBArchivi.FormatoSQL(nomeConfigurazione, DB_TEXT)
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strQuery)
    strValore = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "Valore", valoreDefault)
    
END_GetCtConfigurazione:
    On Local Error GoTo 0
    GetCtConfigurazione = strValore
    
    If Not (HSS Is Nothing) Then
        Call MXDB.dbChiudiSS(HSS)
        Set HSS = Nothing
    End If
    Exit Function
    
ERR_GetCtConfigurazione:
    Dim lngErrNum As Long
    Dim strErrDsc As String
    lngErrNum = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    strValore = valoreDefault
    MXNU.MsgBoxEX "Errore in GetCtConfigurazione. Nome:" & nomeConfigurazione & vbNewLine & "[" & lngErrNum & "] " & strErrDsc, vbCritical, "SisCtLib.cIndCT"
    Resume END_GetCtConfigurazione
    Resume
End Function

Function GetStatoIndicatore(vntNewValore As Variant) As String
Dim strStatoIndicatore As String
Dim strQuery As String

Dim HSS As MXKit.CRecordSet

    strStatoIndicatore = "B"
    On Local Error GoTo ERR_GetStatoIndicatore

    strQuery = "SELECT StatoIndicatore FROM ANAGRAFICACF WHERE CODCONTO = " + hndDBArchivi.FormatoSQL(vntNewValore, DB_TEXT)
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strQuery)
    strStatoIndicatore = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "StatoIndicatore", "B")
    
END_GetStatoIndicatore:
    On Local Error GoTo 0
    GetStatoIndicatore = strStatoIndicatore
    
    If Not (HSS Is Nothing) Then
        Call MXDB.dbChiudiSS(HSS)
        Set HSS = Nothing
    End If
    Exit Function
    
ERR_GetStatoIndicatore:
    Dim lngErrNum As Long
    Dim strErrDsc As String
    lngErrNum = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    strStatoIndicatore = "B"
    MXNU.MsgBoxEX "Errore in GetStatoIndicatore" & vbNewLine & "[" & lngErrNum & "] " & strErrDsc, vbCritical, "SisCtLib.cIndCT"
    Resume END_GetStatoIndicatore
    Resume
End Function

Sub ImpostaColoreCliente(vntNewValore As Variant, cmdct As Object, Optional CodCFFatt As Variant = "")
    On Local Error GoTo ERR_ImpostaColoreCliente
    
    Dim strColore As String
    Dim mCodCli As String
    Dim strStatoIndicatore As String
    
    If CodCFFatt = "" Then
        mCodCli = GetClienteFatturazione(vntNewValore)
    Else
        mCodCli = CodCFFatt
    End If
    strStatoIndicatore = GetStatoIndicatore(mCodCli)

    strColore = GetColoreIndicatore(strStatoIndicatore)
    If strColore = "" Then
        cmdct.Visible = False
        GoTo END_ImpostaColoreCliente
    End If
    
    Dim colori() As String
    colori = Split(strColore, ",")
    
    If UBound(colori) < 2 Then
        cmdct.Visible = False
        GoTo END_ImpostaColoreCliente
    End If
    
    cmdct.BackColor = RGB(colori(0), colori(1), colori(2))
    cmdct.Visible = True

END_ImpostaColoreCliente:
    On Local Error GoTo 0
    Exit Sub
    
ERR_ImpostaColoreCliente:
    Dim lngErrNum As Long
    Dim strErrDsc As String
    lngErrNum = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    cmdct.Visible = False
    MXNU.MsgBoxEX "Errore in ImpostaColoreCliente" & vbNewLine & "[" & lngErrNum & "] " & strErrDsc, vbCritical, "SisCtLib.cIndCT"
    Resume END_ImpostaColoreCliente
    Resume
End Sub

Private Function GetColoreIndicatore(ByVal strStatoIndicatore As String) As String
Dim strRGB As String
Dim strQuery As String

Dim HSS As MXKit.CRecordSet

    strRGB = ""
    On Local Error GoTo ERR_GetColoreIndicatore
    
    strQuery = "SELECT RGB FROM zzTabStatiIndicatore" & _
        " WHERE StatoIndicatore=" & hndDBArchivi.FormatoSQL(strStatoIndicatore, DB_TEXT)
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strQuery)
    strRGB = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "RGB", "")
    
END_GetColoreIndicatore:
    On Local Error GoTo 0
    GetColoreIndicatore = strRGB
    
    If Not (HSS Is Nothing) Then
        Call MXDB.dbChiudiSS(HSS)
        Set HSS = Nothing
    End If
    Exit Function
    
ERR_GetColoreIndicatore:
    Dim lngErrNum As Long
    Dim strErrDsc As String
    lngErrNum = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    strRGB = ""
    MXNU.MsgBoxEX "Errore in GetColoreIndicatore" & vbNewLine & "[" & lngErrNum & "] " & strErrDsc, vbCritical, "SisCtLib.cIndCT"
    Resume END_GetColoreIndicatore
    Resume
End Function

Private Function GetClienteFatturazione(vntNewValore As Variant) As String
Dim strCliFatt As String
Dim strQuery As String

Dim HSS As MXKit.CRecordSet

    strCliFatt = vntNewValore
    On Local Error GoTo ERR_GetClienteFatturazione
    
    strQuery = "SELECT CodContoFatt FROM ANAGRAFICARISERVATICF" & _
        " WHERE Esercizio=" & MXNU.AnnoAttivo & " AND CodConto=" & hndDBArchivi.FormatoSQL(vntNewValore, DB_TEXT)
    Set HSS = MXDB.dbCreaSS(hndDBArchivi, strQuery)
    strCliFatt = MXDB.dbGetCampo(HSS, TIPO_SNAPSHOT, "CodContoFatt", vntNewValore)
    If strCliFatt = "" Then strCliFatt = vntNewValore
    
END_GetClienteFatturazione:
    On Local Error GoTo 0
    GetClienteFatturazione = strCliFatt
    
    If Not (HSS Is Nothing) Then
        Call MXDB.dbChiudiSS(HSS)
        Set HSS = Nothing
    End If
    Exit Function
    
ERR_GetClienteFatturazione:
    Dim lngErrNum As Long
    Dim strErrDsc As String
    lngErrNum = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    strCliFatt = vntNewValore
    MXNU.MsgBoxEX "Errore in GetClienteFatturazione" & vbNewLine & "[" & lngErrNum & "] " & strErrDsc, vbCritical, "SisCtLib.cIndCT"
    Resume END_GetClienteFatturazione
    Resume
End Function

