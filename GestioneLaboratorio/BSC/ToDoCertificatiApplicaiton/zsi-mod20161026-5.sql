IF OBJECT_ID ('dbo.EXCEL_PORTAFOGLIOORDINI') IS NOT NULL
	DROP VIEW dbo.EXCEL_PORTAFOGLIOORDINI
GO

create view EXCEL_PORTAFOGLIOORDINI as 
SELECT 
        R.DATACONSEGNA
        , (CASE WHEN datediff(d, getdate(), R.DATACONSEGNA) < -7 THEN -7 ELSE datediff(d, getdate(), R.DATACONSEGNA) END) AS gg_dataconsegna 
        , (CASE WHEN ACF.STATOBOLLE = 2 THEN '[BLOCCATO] ' ELSE '' END) + AC.DSCCONTO1 AS RAGIONESOCIALE
        , T.TIPODOC + '/' + CAST(T.ESERCIZIO AS VARCHAR) + '/' + CAST(T.NUMERODOC AS VARCHAR) AS DOCUMENTO -- + '/' + CAST(R.POSIZIONE AS VARCHAR) AS DOCUMENTO 
        , T.NUMRIFDOC, ET.RIFCLI 
        , R.CODART AS ARTICOLO
        , R.DESCRIZIONEART AS DESCRIZIONE
 		, I.DESCRIZIONE IMBALLO 
        , CAST(R.QTAGESTRES AS INT) AS QTAGESTRES
        , CAST(isnull(R.NRPEZZIIMBALLO, 0) AS INT) AS NRPEZZIIMBALLO
        , CAST(
        		(CASE WHEN ISNULL(R.NRPEZZIIMBALLO,0) = 0 THEN 0 
        				ELSE ceiling(R.QTAGESTRES/R.NRPEZZIIMBALLO) 
        				END)
         AS INT) FUSTI 
        , isnull(ER.NRLOTTO, 'NESSUN LOTTO') AS NRLOTTO
        , SD.RAGSOCSPED AS SPEDIZIONIERE
        , (CASE WHEN ISNULL(ER.NOTECLI,'') <> '' THEN ER.NOTECLI ELSE NC.NOTE END) NOTECLI 
        , (CASE WHEN ISNULL(ER.NOTEART,'') <> '' THEN ER.NOTEART ELSE NA.NOTE END) NOTEART 
        , ER.NOTEMAG, '' MAGAZZINO 
        , ER.POSIZIONAMENTO
        , isnull(ER.CONFEZIONATO, 0) AS CONFEZIONATO
        , isnull((SELECT TOP 1 x.disponibilita FROM ZSI_DISPONIBILITAM05 x WHERE x.CODCLIFOR = t.codclifor AND x.CODART = r.codart AND x.CODIMBALLO = R.CODIMBALLO), 0) AS DISPONIBILITA 
        , isnull(z.QTACONFEZIONE, isnull(ER.QTASPEDIBILE, 0)) AS QTASPEDIBILE
        , CAST(isnull((CASE WHEN CAST(R.NRPEZZIIMBALLO AS INT) > 0 THEN R.QTAGESTRES/CAST(R.NRPEZZIIMBALLO AS INT) ELSE 0 END), 0) AS INT) AS COLLIDASPEDIRE
        , CAST(isnull((CASE WHEN CAST(R.NRPEZZIIMBALLO AS INT) > 0 THEN ER.QTASPEDIBILE/CAST(R.NRPEZZIIMBALLO AS INT) ELSE 0 END), 0) AS INT) AS COLLISPEDIBILI

        , '' AS TIPOPALLET
        , '' AS ETICHETTATURA
        , T.ESERCIZIO, T.TIPODOC, T.NUMERODOC, T.BIS, T.DATADOC 
        , 'ALFREDO.DEANGELO@GMAIL.COM;knosmailservice@gmail.com' AS EMAIL_CLIENTE --s.ferrigato@zschimmer-schwarz.com;m.michieletti@zschimmer-schwarz.com;
        , T.PROGRESSIVO AS IDTESTA, R.IDRIGA 
        , AA.GRUPPO, AA.CATEGORIA, AA.CODCATEGORIASTAT 
        , T.CODCLIFOR
        , T.DATARIFDOC
        , I.CODICE CODIMBALLO
        , isnull(SD.CODSPED, 0) AS CODSPED
        , R.UMGEST
        , isnull(ER.DATACARICO, R.DATACONSEGNA) AS DATACARICO
        , ISNULL(ER.MAGSTATORIGA, 0) AS MAGSTATORIGA
        , ISNULL(ER.MAGSTATOM05, 0) AS MAGSTATOM05
        , (CASE WHEN ISNULL(ER.MAGSTATORIGA, 0) = 0 THEN '-'
	        WHEN ISNULL(ER.MAGSTATORIGA, 0) = 1 THEN 'INTERAMENTE'
	        WHEN ISNULL(ER.MAGSTATORIGA, 0) = 2 THEN 'PARZIALMENTE'
	        WHEN ISNULL(ER.MAGSTATORIGA, 0) = 3 THEN 'NON SPEDIBILE'
	        ELSE '-' END) AS DSCMAGSTATORIGA
        , (CASE WHEN ISNULL(ER.MAGSTATOM05, 0) = 0 THEN 'INSERITO'
	        WHEN ER.MAGSTATOM05 = 1 THEN 'CONTR. DISP.'
	        WHEN ER.MAGSTATOM05 = 2 THEN 'MANDATO A TRASP.'
	        WHEN ER.MAGSTATOM05 = 3 THEN 'APPRONTATA SPED'
	        WHEN ER.MAGSTATOM05 = 4 THEN 'ASSEGNATI LOTTI'
	        ELSE '-' END) AS DSCMAGSTATOM05
        , (CASE WHEN isnull(EXTRACLIENTI.CLIENTESPECIALE, 'N') = 'S' THEN 1 ELSE 0 END) AS CLIENTESPECIALE
        , T.NUMDESTDIVERSAMERCI
        , R.ANNOTAZIONI
        , ET.NOTECONTAINER
        , (CASE WHEN T.RAGSOCDDM IS NULL THEN 
        		AC.DSCCONTO1 + ' ' + AC.INDIRIZZO + ' ' + AC.LOCALITA + ' ' + AC.PROVINCIA 
        	ELSE
        		t.RAGSOCDDM + ' ' + t.INDIRIZZODDM + ' ' + t.LOCALITADDM + ' ' + t.PROVINCIADDM 
        	END) AS DESTINAZIONE
        , (CASE WHEN isnull(ET.NOTECONTAINER ,'') <> '' THEN 'M14' 
        			WHEN r.CODIMBALLO IN (100, 101) THEN 'M13' 
        			ELSE 'M12' 
			END) AS MODULO
		, ARTADR.NUMONU
        , ARTADR.CLASSEADR
        , '''' + CAST(DAY(ER.DATACARICO) AS VARCHAR) + '/' + CAST(MONTH(ER.DATACARICO) AS VARCHAR) + '/' + CAST(YEAR(ER.DATACARICO) AS VARCHAR) AS XLS_DATACARICO
        , '''' + CAST(DAY(R.DATACONSEGNA) AS VARCHAR) + '/' + CAST(MONTH(R.DATACONSEGNA) AS VARCHAR) + '/' + CAST(YEAR(R.DATACONSEGNA) AS VARCHAR) AS XLS_DATACONSEGNA
        , AC.DSCCONTO1 AS XLS_CLIENTE
        , (CASE WHEN ISNULL(T.RAGSOCDDM, '') = '' THEN 
        		AC.DSCCONTO1 + ' ' + AC.INDIRIZZO + ' ' + AC.LOCALITA + ' ' + AC.PROVINCIA 
        	ELSE
        		t.RAGSOCDDM + ' ' + t.INDIRIZZODDM + ' ' + t.LOCALITADDM + ' ' + t.PROVINCIADDM 
        	END) AS XLS_DESTINAZIONE
        , (CAST(CAST(isnull((CASE WHEN CAST(R.NRPEZZIIMBALLO AS INT) > 0 THEN ER.QTASPEDIBILE/CAST(R.NRPEZZIIMBALLO AS INT) ELSE 0 END), 0) AS INT) AS VARCHAR) + ' ' + I.DESCRIZIONE) AS XLS_TIPOIMBALLO
        , isnull(ER.NRLOTTO, 'NESSUN LOTTO') AS XLS_NRLOTTO
        , ISNULL(ARTADR.NUMONU, '') AS XLS_NUMONU
        , ISNULL(ARTADR.CLASSEADR, 'NO') AS XLS_CLASSEADR
        , ISNULL((SELECT TOP 1 TABSPEDIZ.RAGIONESOCIALE FROM  TABSPEDIZ WHERE TABSPEDIZ.CODICE = isnull(SD.CODSPED, 0)), '') AS XLS_SPEDIZIONIERE
        , R.ANNOTAZIONI AS XLS_NOTE
        , (CASE WHEN ISNULL(ER.NOTEART,'') <> '' THEN ER.NOTEART ELSE NA.NOTE END) XLS_NOTEART 
        , ER.NOTEMAG AS XLS_NOTEMAG
        , 0 AS XLS_COSTO
        , ISNULL(NC.NOTECOSIGNEE, '') AS XLS_NOTECONSIGNEE
        , ISNULL(NC.NOTENOTIFY, '') AS XLS_NOTENOTIFY
        , CAST(CAST((CASE WHEN ISNULL(ER.QTASPEDIBILE, 0) = 0 THEN R.QTAGESTRES ELSE ER.QTASPEDIBILE END) AS INT) AS VARCHAR)AS XLS_QTA
        , ISNULL(ER.DATAPOSIZIONAMENTO, '21000101') AS DATAPOSIZIONAMENTO
        , row_number() OVER(PARTITION BY ER.DATAPOSIZIONAMENTO ORDER BY ER.DATAPOSIZIONAMENTO) AS RN
        , K.IdPubblicazioneKnos AS IDOBJECT
        
FROM TESTEDOCUMENTI T WITH (NOLOCK) JOIN EXTRATESTEDOC ET WITH (NOLOCK) ON T.PROGRESSIVO = ET.IDTESTA 
JOIN RIGHEDOCUMENTI R WITH (NOLOCK) ON T.PROGRESSIVO = R.IDTESTA 
JOIN ANAGRAFICACF AC WITH (NOLOCK) ON AC.CODCONTO = T.CODCLIFOR 
JOIN EXTRACLIENTI WITH (NOLOCK) ON EXTRACLIENTI.CODCONTO = AC.CODCONTO
JOIN ANAGRAFICAARTICOLI AA WITH (NOLOCK) ON AA.CODICE = R.CODART 
JOIN EXTRARIGHEDOC ER WITH (NOLOCK) ON ER.IDTESTA = R.IDTESTA AND ER.IDRIGA = R.IDRIGA 
JOIN ANAGRAFICARISERVATICF ACF WITH (NOLOCK) ON ACF.ESERCIZIO = YEAR(GETDATE()) AND ACF.CODCONTO = T.CODCLIFOR 
LEFT OUTER JOIN ZSI_CONFEZSPEDIBILE z WITH (NOLOCK) ON Z.PROGRESSIVO = r.IDTESTA AND z.idriga = r.IDRIGA
LEFT OUTER JOIN TAB_ARTICOLIADR ARTADR WITH (NOLOCK) ON ARTADR.CODART = R.CODART
LEFT OUTER JOIN TABIMBALLI I WITH (NOLOCK) ON R.CODIMBALLO = I.CODICE 
LEFT OUTER JOIN ZSI_NOTECLI NC WITH (NOLOCK) ON T.CODCLIFOR = NC.CODCLI AND (T.NUMDESTDIVERSAMERCI = NC.CODDEST OR (T.NUMDESTDIVERSAMERCI = 0 AND NC.CODDEST = -1) OR (T.NUMDESTDIVERSAMERCI <> 0 AND NOT EXISTS(SELECT 1 FROM ZSI_NOTECLI Z WHERE Z.CODCLI = T.CODCLIFOR AND Z.CODDEST = T.NUMDESTDIVERSAMERCI) AND NC.CODDEST = -1) ) 
LEFT OUTER JOIN ZSI_NOTEART NA WITH (NOLOCK) ON T.CODCLIFOR = NA.CODCLI AND (T.NUMDESTDIVERSAMERCI = NA.CODDEST OR (T.NUMDESTDIVERSAMERCI = 0 AND NC.CODDEST = -1) OR (T.NUMDESTDIVERSAMERCI <> 0 AND NOT EXISTS(SELECT 1 FROM ZSI_NOTECLI Z WHERE Z.CODCLI = T.CODCLIFOR AND Z.CODDEST = T.NUMDESTDIVERSAMERCI) AND NC.CODDEST = -1) ) AND R.CODART = NA.CODART 
LEFT OUTER JOIN SPEDIZDOCUMENTI SD WITH (NOLOCK) ON SD.IDTESTA = t.PROGRESSIVO
LEFT OUTER JOIN ITA_ARCHIVIO_DOCUMENTI K ON K.Progressivo = T.PROGRESSIVO
WHERE 
	t.tipodoc IN ('OCC', 'OCE') AND 
	R.CODART <> '' AND
	R.QTAGESTRES > 0 
	AND t.esercizio > year(getdate())-1
GO

GRANT DELETE ON dbo.EXCEL_PORTAFOGLIOORDINI TO Metodo98
GO
GRANT INSERT ON dbo.EXCEL_PORTAFOGLIOORDINI TO Metodo98
GO
GRANT REFERENCES ON dbo.EXCEL_PORTAFOGLIOORDINI TO Metodo98
GO
GRANT SELECT ON dbo.EXCEL_PORTAFOGLIOORDINI TO Metodo98
GO
GRANT UPDATE ON dbo.EXCEL_PORTAFOGLIOORDINI TO Metodo98
GO

