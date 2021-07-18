Attribute VB_Name = "MCostanti"

' nome degli oggetti plug-in
Global Const OGGETTO_ESTENSIONE = "objExt"
Global Const OGGETTO_WRAPPER_ESTENSIONE = "objExtWrapper"

'DATABASE
Global Const NOME_DB_ABICAB = "abicab"
'Global Const NOME_DB_DITTE = "metditte"
'Global Const NOME_DB_OGG = "metogg"
Global Const NOME_DB_ARCHIVI = "metxxxx"

Global Const NOME_FILE_MENU_TMP = "menu.ini"
Global Const NOME_FILE_MENUTOOLS_TMP = "menutools.ini"
Global Const NOME_FILE_MENUPERS_TMP = "menup.ini"
Global Const NOME_FILE_MENUPERSDITTA_TMP = "menud.ini"

'errori
Global Const ERR_UTENTE = 3059
Global Const ERR_DBVARIAZIONE = 20000



Global Const KEY_F1 = &H53    '&H70 -> sostituito con S
Global Const HOURGLASS = 11     ' 11 - Hourglass

Global Const SHIFT_MASK = 1
Global Const CTRL_MASK = 2
Global Const ALT_MASK = 4

Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed

Global Const OLE_ACTIVATE = 7

'Common Dialog Control
'Action Property
Global Const DLG_FILE_OPEN = 1
Global Const DLG_FILE_SAVE = 2
Global Const DLG_COLOR = 3
Global Const DLG_FONT = 4
Global Const DLG_PRINT = 5
Global Const DLG_HELP = 6

'tasti corrispondenti ai bottoni della toolbox
Global Const BTN_INS = 73 'I
Global Const BTN_MOD = 68 '"D"
Global Const BTN_REG = 82 '"R"
Global Const BTN_PRIMO = 49 '"1"
Global Const BTN_PREC = 50 '"2"
Global Const BTN_SUCC = 51 '"3"
Global Const BTN_ULTIMO = 52 '"4"
Global Const BTN_ANN = 65 '"A"
Global Const BTN_DEF_ACC = 85 '"U"
Global Const BTN_STP = 80 '"P"
Global Const BTN_DUP = 48 '"0"
Global Const BTN_VISUTMOD = 77 '"M"
Global Const BTN_ALLINONE = 78 '"N"
Global Const BTN_ZOOM = 90 '"N"
Global Const BTN_TROVA = 84 '"T"
Global Const BTN_DESIGNER = 87 '"W"


'maschere dei tasti
Global Const BTN_INS_MASK = &H1
Global Const BTN_MOD_MASK = &H2
Global Const BTN_REG_MASK = &H4
Global Const BTN_PRIMO_MASK = &H8
Global Const BTN_PREC_MASK = &H10
Global Const BTN_SUCC_MASK = &H20
Global Const BTN_ULTIMO_MASK = &H40
Global Const BTN_ANN_MASK = &H80
Global Const BTN_TUTTI_MASK = &H1 + &H2 + &H4 + &H8 + &H10 + &H20 + &H40 + &H80
Global Const BTN_STP_MASK = &H100

Global Const SQL_SUCCESS As Long = 0
Global Const SQL_FETCH_NEXT As Long = 1
Global Const ODBC_ADD_SYS_DSN = 4
Global Const ODBC_REMOVE_SYS_DSN = 6

'formule preventivo commessa
Global MobjScriptFormuleCC As Object

'--------- costanti gestione movimenti ------------------------
Public Enum setGestioneMovimentiAzione
    REC_ANNULLA = 1
    REC_INSERISCI = 2
    REC_MODIFICA = 3
End Enum

'Public Enum setGestioneMovimentiOperazione
'    MOV_AGGIORNA = 1
'    MOV_STORNO = -1
'End Enum

'costanti per TIPOMOV:
'Public Enum setStoricoTipoMovimento
'    ST_MOV_MANUALE = 0
'    ST_MOV_RIGADOC = 1
'    ST_MOV_RIGADOC_COLL = 2
'    ST_MOV_COMP = 3
'    ST_MOV_COMP_COLL = 4
'    ST_MOV_COMPCOMM = 5           'Componenti Commessa Prod.
'    ST_MOV_COMPCOMM_COLL = 6      'Componenti Commessa Prod. Collegati
'End Enum

'*** DEFINIZIONI TIPI DI DATI ***
'*** METODO.BAS ***
Type Proprieta_Aggiuntive
     'Tipo As String * 1 'vedere sotto le costanti
     Tipo As setTipoInput
     DataF As String
     frmt As String
     dflt As Variant
     ValCorrente As Variant 'per gestire le modifiche del campo
End Type

Type SS_Prop_Aggiuntive
    Row As Long
    Col As Long
    Tipo As String * 1 'vedere sotto le costanti
    DataF As String
    'frmt As String
    dflt As Variant
    'ValCorrente As Variant 'per gestire le modifiche del campo
End Type

'*** MAGAZZIN.BAS ***
Type Dati_tipologia
    cod As String * 2     'da TabTipologie e da TipologieArticoli
    des As String * 25
    CTRLEs As Integer
    lngvar As Integer
    SelVar As String * 1
    aggDes As Integer
    varcar As Integer
    nr As Integer
    hsnapvar As Integer 'indice dello snapshot delle varianti usato solo per la generazione automatica
End Type

Type StrDisponibilita
    Giacenze1UM(1 To 10) As Currency    'Giacenze Prima Unità di Misura
    TotGiacenze1UM As Currency          'Totale Giacenze 1UM
    Giacenze2UM(1 To 10) As Currency    'Giacenze Seconda Unità di Misura
    TotGiacenze2UM As Currency          'Totale Giacenze 2UM
    Ordinato(1 To 2) As Currency        'Ordinato 1UM e 2UM
    Impegnato(1 To 2) As Currency       'Impegnato 1UM e 2UM
    GiacenzaIniziale(1 To 2) As Currency 'Giacenza Iniziale 1UM e 2UM
End Type

'Struttura per l'aggiornamento Prezzo/Sconto nella tabella GestionePrezzi
Type ParPrezziSconto
    ProgV As Long
    ProgN As Long
    CliFor As String
    CodArt As String
    Listino As Integer
    prez As Variant
    DataInizioVal As String
    TipoCampo As Integer  '1=Prezzo 2=Sconto
End Type

'Struttura per la ricerca del Magazzino
Type DatiRicMag
    TipoRicerca As Integer     '0=Ricerca su Parametri Doc; 1=Ricerca su Parametri Ord. Prod.
    CodiceDoc As String
    CodiceArt As String
    CodConto As String
    NumDestDiv As Integer
    TipoMag As Integer
    CodMagRP As String   'Codice Magazzino della Riga Prodotto
End Type

'struttura per la stampa differita documenti
Type Parametri_Stampa_Documento
    NrTerminale As Integer
    AnnoDoc As Integer '(rif 10)
    TipoDoc As String
    NumeroDoc As Long
    bis As String
    DataDoc As String
    CodConto As String
    Lingua As Integer
    StampaVar As String
    'StampaDscLingua As Integer
    OpzioniStampa As Integer
    StampaDistBase As String
    SaltoPag    As Integer
    StampaInfo  As Integer
    DEVStampa As String
    DEVStampaInfo As String
    ModuloStampaDist As String
    DEVStampaEtic As String
    ModuloStampaEtic As String
    TipoStampaEtic As Integer
End Type


'       COSTANTI
Global Const LISTA_TABELLE_STD = 0
Global Const LISTA_VALIDAZIONI_STD = 1
Global Const LISTA_VISIONI_STD = 2
Global Const LISTA_SITUAZIONI_STD = 3
Global Const LISTA_ANAGRAFICHE_STD = 4
Global Const LISTA_MULTIANAGRAFICHE_STD = 5
Global Const LISTA_TABELLE_PERS = 6
Global Const LISTA_VALIDAZIONI_PERS = 7
Global Const LISTA_VISIONI_PERS = 8
Global Const LISTA_SITUAZIONI_PERS = 9
Global Const LISTA_ANAGRAFICHE_PERS = 10
Global Const LISTA_MULTIANAGRAFICHE_PERS = 11
Global Const LISTA_TABELLE_PERSDITTA = 12
Global Const LISTA_VALIDAZIONI_PERSDITTA = 13
Global Const LISTA_VISIONI_PERSDITTA = 14
Global Const LISTA_SITUAZIONI_PERSDITTA = 15
Global Const LISTA_ANAGRAFICHE_PERSDITTA = 16
Global Const LISTA_MULTIANAGRAFICHE_PERSDITTA = 17
Global Const LISTA_RISORSE_TOOLBAR = 18
Global Const LISTA_RISORSE_MSGBOX = 19
Global Const LISTA_RISORSE_ETICHETTE = 20
Global Const LISTA_RISORSE_LINGUETTE = 21
Global Const LISTA_RISORSE_CAPTIONFORM = 22
Global Const LISTA_RISORSE_VARIE = 23
Global Const LISTA_RISORSE_BOTTONI = 24
Global Const LISTA_RISORSE_FOGLI = 25
Global Const LISTA_RISORSE_POPUP = 26
Global Const LISTA_RISORSE_CHECK = 27
Global Const LISTA_RISORSE_OPTION = 28
Global Const LISTA_RISORSE_STATUS = 29
Global Const LISTA_RISORSE_COMBO = 30
Global Const LISTA_RISORSE_ERRORI = 31
Global Const LISTA_RISORSE_TITOLO_VIS = 32
Global Const LISTA_RISORSE_TIPO_VIS = 33
Global Const LISTA_RISORSE_TITOLO_SIT = 34
Global Const LISTA_RISORSE_FILTROVEL = 35
Global Const LISTA_RISORSE_TITOLOTOT = 36
Global Const LISTA_RISORSE_NOMECOL = 37
Global Const LISTA_RISORSE_STAMPE = 38
Global Const LISTA_RISORSE_TEMA = 39
Global Const LISTA_RISORSE_ITALCOM = 40
Global Const ODBCDAT = 41

Global Const PROCESS_QUERY_INFORMATION = &H400

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F



