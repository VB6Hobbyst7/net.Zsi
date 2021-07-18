Attribute VB_Name = "MVarieExt"
Option Explicit
DefLng A-Z

Global GIstanze As Long
#If TOOLS <> 1 Then
    Public Enum enmTestSalva
        tsNessuno = 0
        tsSalvato = 1
        tsNonSalvato = 2
        tsRitorna = 3
    End Enum
#End If

Private Const WM_CLOSE = &H10
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Sub UnloadFormExt(frm As Form)
    Dim lngRes&
    Dim Hwndfrm&
    
    Hwndfrm = frm.hwnd
    lngRes = PostMessage(Hwndfrm, WM_CLOSE, 0&, 0&)

End Sub


Public Function Inizializza_i(colAmbienti As Collection, colOggettiGlobali As Collection) As Boolean

    Dim vntobj As Variant
    
    On Local Error GoTo err_Inizializza_i
    
    Inizializza_i = False
    If GIstanze = 0 Then
    
         For Each vntobj In colAmbienti
             Select Case UCase$(TypeName(vntobj))
                 Case "XNUCLEO"
                     'Set MXNU = Nothing
                     Set MXNU = vntobj
                 Case "XODBC"
                     Set MXDB = vntobj
                 Case "CAMBCRW"
                     Set MXCREP = vntobj
                 Case "CAMBAGENTI"
                     Set MXAA = vntobj
                 Case "CAMBTAB"
                     Set MXCT = vntobj
                 Case "CAMBVISIONI"
                     Set MXVI = vntobj
                 Case "CAMBVALID"
                     Set MXVA = vntobj
                 Case "CAMBFILTRI"
                     Set MXFT = vntobj
                 Case "CAMBSCAD"
                     Set MXSC = vntobj
                 Case "CAMBVART"
                     Set MXART = vntobj
                 Case "CAMBSTMAG"
                     Set MXSM = vntobj
                 Case "CAMBDBA"
                     Set MXDBA = vntobj
                 Case "CAMBGESTDOC"
                     Set MXGD = vntobj
                 Case "CAMBPIAN"
                     Set MXPIAN = vntobj
                 Case "CAMBPN"
                     Set MXPN = vntobj
                 Case "CAMBPROD"
                     Set MXPROD = vntobj
                 Case "CAMBCOMMCLI"
                     Set MXCC = vntobj
                 Case "CAMBCICLILAV"
                     Set MXCICLI = vntobj
                 Case "CAMBRISORSE"
                     Set MXRIS = vntobj
             End Select
         Next
        
         For Each vntobj In colOggettiGlobali
             Select Case UCase$(TypeName(vntobj))
                 Case "CCONNESSIONE"
                     Set hndDBArchivi = vntobj
             End Select
         Next
    
    End If
    GIstanze = GIstanze + 1
    Inizializza_i = True
    On Local Error GoTo 0
fine_Inizializza_i:
    
Exit Function

err_Inizializza_i:
    MsgBox "Errore [" & Err.Number & " - " & Err.Description & "] durante l'inizializzazione dell'estensione", vbCritical, "Attenzione!"
    Resume fine_Inizializza_i
End Function


Public Sub Termina_i()

    GIstanze = GIstanze - 1
    
    If GIstanze = 0 Then
        Set hndDBArchivi = Nothing
        
        Set MXRIS = Nothing
        Set MXCC = Nothing
        Set MXCICLI = Nothing
        Set MXPROD = Nothing
        Set MXPIAN = Nothing
        Set MXGD = Nothing
        Set MXPN = Nothing
        Set MXSM = Nothing
        Set MXDBA = Nothing
        Set MXART = Nothing
        Set MXCT = Nothing
        Set MXSC = Nothing
        Set MXVA = Nothing
        Set MXAA = Nothing
        Set MXVI = Nothing
        Set MXFT = Nothing
        Set MXCREP = Nothing
        Set MXDB = Nothing
        Set MXNU = Nothing
    End If

End Sub

Public Sub DisabilitaControlliExt(colControls As Object)

    Dim objControl As Control
    
    For Each objControl In colControls
        If TypeName(objControl) = "TextBox" Then
            objControl.Enabled = False
        End If
    Next
    DoEvents

End Sub
