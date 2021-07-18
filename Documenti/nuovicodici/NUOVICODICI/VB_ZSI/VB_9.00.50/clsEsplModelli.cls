VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEsplModelli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mvarNomeImpostazione As String
Dim mvarFldCampoPerc As String


Public Property Let NomeImpostazione(ByVal v As String)
    mvarNomeImpostazione = v
End Property

Public Property Get NomeImpostazione() As String
    NomeImpostazione = mvarNomeImpostazione
End Property

Public Property Let CampoPerc(ByVal v As String)
    mvarFldCampoPerc = v
End Property

Public Property Get CampoPerc() As String
    CampoPerc = mvarFldCampoPerc
End Property

Public Sub Esplodi(ByVal CodModello As String, ByVal Qta As Double, Arr_Modelli() As Variant)



'
'Input:
'       CodModello:  'modello' fino alla tipologia classe inclusa o 'gruppo' di modelli (tabella_gruppi)
'       Qta:         quantita' richiesta sul modello indicato in CodModello
'       Arr_Modelli: Array di variant (bidimensionale, allocato da questa procedura Sub)
'
'OutPut:
'       L' array caricato con le informazioni indicate in "Descrizione"
'
'
'Descrizione: Calcola le percentuali di qta relative al modello in esame
'             sulla base della quantita' fornita in input.
'             Se in in input viene fornito il modello,
'             le quantita' sono semplicemente date dalla percentuale x il modello.
'             Viceversa, sono rapportate ad un gruppo di modelli
'             considerati come un unico modello, ovvero
'             la cui somma delle quantita' splittate rappresenta il valore passato
'             in ingresso.
'
'>> Restituisce un vettore bidimensionale cosi' composto:
'
'   Modello     Grafica     Taglia      Perc %                  <Qta calcolata>
'                                       (all' interno del modello)
'   056#100     04000       1           10.2                                102
'   021#100     06000       2           5                                    50
'
'   (Supponendo che siano state richieste le qta "splittate" per una quantita'
'   complessiva di 1000 pezzi)




'
'Verifica se il modello specificato e' un modello o un gruppo di modelli.
'


Dim strSQL As String

strSQL = "SELECT * FROM TAB_GRUPPIMODELLI INNER JOIN (SELECT * FROM TAB_SPLIT WHERE " & _
         "NOMEIMPOSTAZIONE = ':NOMEIMPOSTAZIONE') TB " & _
         "ON TAB_GRUPPIMODELLI.MODELLO = TB.MODELLO  WHERE CODGRUPPO = ':CODGRUPPO'"
         
strSQL = Replace(strSQL, ":CODGRUPPO", CodModello)
strSQL = Replace(strSQL, ":NOMEIMPOSTAZIONE", mvarNomeImpostazione)

Dim HDY As CRecordSet
Dim bFlgGruppo As Boolean
Dim TotPerc As Double
Dim Perc As Double
Dim Modello As String
Dim Grafica As String
Dim Taglia As String
Dim rccount As Long
Dim j As Long

j = 0

TotPerc = 0

With MXDB
    Set HDY = .dbCreaDY(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
    If Not .dbFineTab(HDY, TIPO_DYNASET) Then
        '
        'E' un codice di gruppo di modelli.
        '
        bFlgGruppo = True
    Else
        .dbChiudiDY HDY
        strSQL = "SELECT * FROM TAB_SPLIT WHERE NOMEIMPOSTAZIONE = ':NOMEIMPOSTAZIONE' AND MODELLO = ':MODELLO' "

        strSQL = Replace(strSQL, ":MODELLO", CodModello)
        strSQL = Replace(strSQL, ":NOMEIMPOSTAZIONE", mvarNomeImpostazione)
        Set HDY = .dbCreaDY(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
        bFlgGruppo = False
    End If
    
    If Not .dbFineTab(HDY) Then
        rccount = HDY.RecSet.RecordCount
        ReDim Arr_Modelli(0 To rccount - 1, 0 To 4) As Variant
    End If

    Do While Not .dbFineTab(HDY)
        
        Perc = .dbGetCampo(HDY, TIPO_SNAPSHOT, mvarFldCampoPerc, 0)
        
        If bFlgGruppo Then              ' se stiamo lavorando sui gruppi totalizza le percentuali
            TotPerc = TotPerc + Perc
        End If
        
        Modello = .dbGetCampo(HDY, TIPO_SNAPSHOT, "MODELLO", "")
        Grafica = .dbGetCampo(HDY, TIPO_SNAPSHOT, "GRAFICA", "")
        Taglia = .dbGetCampo(HDY, TIPO_SNAPSHOT, "TAGLIA", "")
        
        Arr_Modelli(j, 0) = Modello
        Arr_Modelli(j, 1) = Grafica
        Arr_Modelli(j, 2) = Taglia
        Arr_Modelli(j, 3) = Perc
        Arr_Modelli(j, 4) = 0
        .dbSuccessivo HDY, TIPO_SNAPSHOT
        j = j + 1
    Loop
    .dbChiudiDY HDY
    
    Dim i As Long
    
    For i = 0 To UBound(Arr_Modelli(), 1)
        If bFlgGruppo Then  ' se si stanno calcolando le percentuali dei gruppi si divide anche per il totale
                            ' delle percentuali, in modo da tenere conto del fatto che quest' ultime sono riferite
                            ' al gruppo.
            Arr_Modelli(i, 4) = Qta * Arr_Modelli(i, 3) / TotPerc
        Else
            Arr_Modelli(i, 4) = Qta * Arr_Modelli(i, 3)
        End If
    Next i
    
End With
Set HDY = Nothing

End Sub

