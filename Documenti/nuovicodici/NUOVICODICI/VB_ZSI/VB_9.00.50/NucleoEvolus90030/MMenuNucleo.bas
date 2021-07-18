Attribute VB_Name = "MMenuNucleo"
Option Explicit
DefLng A-Z


Public Function EseguiAzione_i(ByVal NomeMenu As Variant, ByVal indice As Integer, HelpContextID As Long)

    Dim q As Long
    Dim agt As String
    Dim par As Variant
    Dim frmGenerica As Form
    Dim bolUscita As Boolean


    'Oggetti necessari alle form general purpose di gestione filtro+tabelle e visioni con selezione
    Dim CEsegui As Object

    bolUscita = False
    On Local Error Resume Next
    frmModuli.Enabled = False
    On Local Error GoTo 0

    If Not EseguiAzioneINI(NomeMenu & "_" & indice, HelpContextID, UCase$(Right(NomeMenu, 4)) = "PERS") Then

        Select Case NomeMenu
    'COMUNI ------------------------------
            Case "MenuItem"
                Select Case indice
                    Case 0
                        If ApriDittaAnno(False, "", "") Then
                            'If Not MXNU.Key_RivBuffetti() Then
                            '    mnuDitteItem(4).Visible = (MXNU.StatoEsercizioCont = 0)
                            'End If
                            'mnuDitteItem(3).Visible = (MXNU.StatoEsercizioCont = 1 Or MXNU.StatoEsercizioCont = 3)
                        End If
                    Case 1
                        Dim NuovoAnno As Integer
                        If SelezioneAnno(True, NuovoAnno) Then
                            MXNU.AnnoAttivo = NuovoAnno
                            Call ChiudiFormAttive
                            Call ApriAnno(True)
                            'If Not MXNU.Key_RivBuffetti() Then
                            '    mnuDitteItem(4).Visible = MXNU.StatoEsercizioCont = 0
                            'End If
                            'mnuDitteItem(3).Visible = (MXNU.StatoEsercizioCont = 1 Or MXNU.StatoEsercizioCont = 3)
                        End If
                    Case 3
                    
                        SetCursorPos 10, ((metodo.Height - metodo.ScaleHeight) + 30) / Screen.TwipsPerPixelY
                        Call AttivaMenuMetodo
                    Case 5
                        Call CambioUtenteAttivo
                    Case 7 'Uscita
                        Unload metodo
                        bolUscita = True
                End Select
            Case "AiutoItem"
                Select Case indice
                    Case 0
                        Call MXNU.ApriHelp(True)
                    Case 2 'informazioni su
                        frmInformazioni.Show 1
                End Select
            #If ISKEY = 1 Then
            '-------------------
            Case "TabelleXItem" ' prog chiavi
                Select Case indice
                    Case 0 'anagrafica rivenditori
                        Call funzione2(HelpContextID)
                    Case 1 'anagrafica clienti
                        Call funzione3(HelpContextID)
                End Select
            Case "TabelleX1Item" ' prog chiavi
                Select Case indice
                    Case 0 'tabella
                End Select
            Case "TabelleX2Item" ' prog chiavi
                Select Case indice
                    Case 0 'anagrafica chiavi
                        Call funzione4(HelpContextID)
                    Case 1 'lettura chiave
                        Load Chiave
                        Chiave.Show vbModal
                End Select
            '---------------------
            #End If
        End Select
    End If

    On Local Error Resume Next
    If Not bolUscita Then frmModuli.Enabled = True
    On Local Error GoTo 0
End Function

