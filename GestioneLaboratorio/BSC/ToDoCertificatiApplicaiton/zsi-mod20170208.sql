IF NOT EXISTS(SELECT 1 FROM INFORMATION_SCHEMA.columns WHERE table_name = 'EXTRATESTEDOC' AND COLUMN_NAME = 'CODSPEDMANDATO')
ALTER TABLE EXTRATESTEDOC ADD CODSPEDMANDATO DECIMAL(5)
GO



IF OBJECT_ID ('dbo.VistaRelazioniCFVdaDocZSIM05') IS NOT NULL
	DROP VIEW dbo.VistaRelazioniCFVdaDocZSIM05
GO

CREATE VIEW VistaRelazioniCFVdaDocZSIM05 AS 
 SELECT
     TD.PROGRESSIVO,
     RD.IDRIGA,
     RCFV.DSCART_ALT
 FROM
     TESTEDOCUMENTI TD INNER JOIN RIGHEDOCUMENTI RD ON TD.PROGRESSIVO = RD.IDTESTA
     LEFT OUTER JOIN RELAZIONICFV RCFV ON RD.CODART = 
          (CASE PATINDEX('%#%', RD.CODART) WHEN 0 THEN RCFV.ARTICOLO ELSE (RCFV.ARTICOLO + '#' + RCFV.VARIANTI) END) 
 AND TD.CODCLIFOR = RCFV.CODCLIFOR  

GO

GRANT SELECT ON dbo.VistaRelazioniCFVdaDocZSIM05 TO Metodo98
GO





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
        , (R.DESCRIZIONEART + ' (' + ISNULL(AA.DESCRIZIONE, '')  + ')') AS DESCRIZIONE
 		, I.DESCRIZIONE IMBALLO 
        , CAST(R.QTAGESTRES AS INT) AS QTAGESTRES
        , CAST(isnull(R.NRPEZZIIMBALLO, 0) AS INT) AS NRPEZZIIMBALLO
        , CAST(
        		(CASE WHEN ISNULL(R.NRPEZZIIMBALLO,0) = 0 THEN 0 
        				ELSE ceiling(R.QTAGESTRES/R.NRPEZZIIMBALLO) 
        				END)
         AS INT) FUSTI 
        , isnull(ER.NRLOTTO, 'NESSUN LOTTO') AS NRLOTTO
        , ISNULL(SD.RAGIONESOCIALE, '') AS SPEDIZIONIERE
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
        , isnull(isnull(ET.CODSPEDMANDATO, ACF.CODSPED), 0) AS CODSPED
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
        --, ISNULL((SELECT TOP 1 TABSPEDIZ.RAGIONESOCIALE FROM  TABSPEDIZ WHERE TABSPEDIZ.CODICE = isnull(SD.CODSPED, 0)), '') AS XLS_SPEDIZIONIERE
        , ISNULL(SD.RAGIONESOCIALE, '') AS XLS_SPEDIZIONIERE
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
LEFT OUTER JOIN ZSI_NOTEART NA WITH (NOLOCK) ON T.CODCLIFOR = NA.CODCLI AND R.CODART = NA.CODART 
	 AND (T.NUMDESTDIVERSAMERCI = NA.CODDEST 
	 OR (T.NUMDESTDIVERSAMERCI = 0 AND NC.CODDEST = -1) 
	 OR (T.NUMDESTDIVERSAMERCI <> 0 AND NOT EXISTS(SELECT 1 FROM ZSI_NOTEART Z WHERE Z.CODCLI = T.CODCLIFOR AND Z.CODDEST = T.NUMDESTDIVERSAMERCI) AND NC.CODDEST = -1) 
	 )
--LEFT OUTER JOIN SPEDIZDOCUMENTI SD WITH (NOLOCK) ON SD.IDTESTA = t.PROGRESSIVO
LEFT OUTER JOIN TABSPEDIZ SD WITH (NOLOCK) ON ET.CODSPEDMANDATO = SD.CODICE
LEFT OUTER JOIN ITA_ARCHIVIO_DOCUMENTI K ON K.Progressivo = T.PROGRESSIVO
WHERE 
	t.tipodoc IN ('OCC', 'OCE') AND 
	R.CODART <> '' AND
	R.QTAGESTRES > 0 
	AND t.esercizio >= year(getdate())-2
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



SELECT TOP 1 * FROM EXCEL_PORTAFOGLIOORDINI

IF OBJECT_ID ('dbo.ITA_SP_UPDATE_DATISPEDIZIONE') IS NOT NULL
	DROP PROCEDURE dbo.ITA_SP_UPDATE_DATISPEDIZIONE
GO

CREATE PROCEDURE [dbo].[ITA_SP_UPDATE_DATISPEDIZIONE]
(

@idtesta DECIMAL(10)
, @idriga DECIMAL(5)
, @codsped AS DECIMAL(5)
, @qtaspedibile AS DECIMAL(16,6)
, @disponibilita AS DECIMAL(16,6)
, @posizionamento AS VARCHAR(50)
, @nrlotto AS VARCHAR(50)
, @magstatoriga AS SMALLINT 
, @magstatom05 SMALLINT
, @notecli VARCHAR(3000)
, @noteart VARCHAR(3000)
, @notemag VARCHAR(3000)
, @annotazioni VARCHAR(100)
, @coddestdiv DECIMAL(5)
, @noteconsignee VARCHAR(1000)
, @notenotify VARCHAR(1000)
, @notecontainer VARCHAR(1000)
, @datacarico DATETIME
, @dataconsegna DATETIME
, @dataposizionamento DATETIME
)
AS

BEGIN

	SET NOCOUNT ON

	DECLARE @codart VARCHAR(50)
	SELECT @codart = CODART FROM RIGHEDOCUMENTI WHERE IDTESTA = @idtesta AND IDRIGA = @idriga

	DECLARE @codcli VARCHAR(7)
	SELECT @codcli = CODCLIFOR, @coddestdiv = (CASE WHEN ISNULL(NUMDESTDIVERSAMERCI, 0) = 0 THEN -1 ELSE NUMDESTDIVERSAMERCI END) 
	FROM TESTEDOCUMENTI WHERE PROGRESSIVO = @idtesta 

	DECLARE @esistenota SMALLINT
	
	UPDATE EXTRATESTEDOC 
	SET notecontainer = @notecontainer
	WHERE IDTESTA = @idtesta
	
	-- aggiornamento dati riga
	UPDATE RIGHEDOCUMENTI
	SET ANNOTAZIONI = @annotazioni
	, DATACONSEGNA = @dataconsegna
	WHERE IDTESTA = @idtesta AND idriga = @idriga
	
	
	
	DECLARE @nrlotti VARCHAR(1000)
	SELECT @nrlotti = dbo.fn_GetLOTTIM05(@idtesta, @idriga)

	IF @nrlotti <> ''
	BEGIN
		SET @nrlotto = @nrlotti
	END
	
	DECLARE @qtalotti AS DECIMAL(16,6)
	SELECT @qtalotti = sum(QTACONFEZIONE) 
	FROM ZSI_CONFEZSPEDIBILE
	WHERE PROGRESSIVO = @idtesta AND idriga = @idriga
	GROUP BY PROGRESSIVO, idriga
	
	IF isnumeric(@qtalotti) = 1
	BEGIN
		SET @qtaspedibile  = @qtalotti	
	end
	
	-- aggiornamento dati extra
	UPDATE EXTRARIGHEDOC
	SET --DISPONIBILITA = @disponibilita	, 
	NRLOTTO = @nrlotto
	, POSIZIONAMENTO = @posizionamento
	--, CONFEZIONATO = @confezionato
	, QTASPEDIBILE = @qtaspedibile
	, MAGSTATOM05 = @magstatom05
	, MAGSTATORIGA = @magstatoriga
	, NOTEMAG = @notemag
	, DATACARICO = @datacarico
	, DATAPOSIZIONAMENTO = @dataposizionamento
	WHERE IDTESTA = @idtesta AND idriga = @idriga
	
	
	EXEC ITA_SP_UPDATE_DISPONIBILITA @idtesta, @idriga, @disponibilita

	
	SELECT @esistenota = count(*) FROM ZSI_NOTEART
	WHERE CODART = @codart AND CODCLI = @codcli AND CODDEST = @coddestdiv
	
	IF (@esistenota = 0)
	BEGIN
		INSERT INTO dbo.ZSI_NOTEART (CODCLI, CODDEST, CODART, NOTE, UTENTEMODIFICA, DATAMODIFICA)
		VALUES (@codcli, @coddestdiv, @codart, @noteart, 'trm', getdate())
	END
	

	UPDATE ZSI_NOTEART
	SET NOTE = @noteart
	WHERE CODART = @codart AND CODCLI = @codcli AND CODDEST = @coddestdiv




	SELECT @esistenota = count(*) FROM ZSI_NOTECLI
	WHERE CODCLI = @codcli AND CODDEST = @coddestdiv
	
	IF (@esistenota = 0)
	BEGIN
		INSERT INTO dbo.ZSI_NOTECLI (CODCLI, NOTE, UTENTEMODIFICA, DATAMODIFICA, CODDEST, NOTECOSIGNEE, NOTENOTIFY)
		VALUES (@codcli, @notecli, 'trm', getdate(), @coddestdiv, @noteconsignee, @notenotify)
	END
		

	UPDATE ZSI_NOTECLI
	SET NOTE = @notecli
	, NOTECOSIGNEE =  @noteconsignee
	, NOTENOTIFY = @notenotify
	WHERE CODCLI = @codcli AND CODDEST = @coddestdiv

	
	-- aggiornamento spedizioniere
	IF @codsped > 0
	BEGIN
	
		UPDATE EXTRATESTEDOC SET CODSPEDMANDATO = @CODSPED WHERE IDTESTA = @IDTESTA
	
--		DELETE FROM SPEDIZDOCUMENTI WHERE IDTESTA = @idtesta
		
--		INSERT INTO dbo.SPEDIZDOCUMENTI (IDTESTA, POSSPED, POSIZIONE, CODSPED, RAGSOCSPED, INDIRIZZOSPED, CAPSPED, LOCALITASPED, PROVSPED, UTENTEMODIFICA, DATAMODIFICA, PARTITAIVA, CODNAZIONE, NUMALBOTR)
--		SELECT @idtesta, 1, 1, @codsped, RAGIONESOCIALE, indirizzo, CAP, LOCALITA, PROVINCIA, 'trm', getdate(), PARTITAIVA, CODNAZIONE, NUMALBOTR
--		FROM TABSPEDIZ WHERE CODICE = @codsped
			
			
	END
	
	
	-- appuntamenti
	DELETE FROM Appointments WHERE Start between @dataposizionamento AND dateadd(d, 1, @dataposizionamento) --mkey = CAST(@idtesta AS VARCHAR) + '|' + CAST(@idriga AS VARCHAR)
	
	INSERT INTO dbo.Appointments (Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey)
	SELECT 
	Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey
	FROM EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS e 
	WHERE e.dataposizionamento = @dataposizionamento -- idtesta = @idtesta AND idriga = @idriga	
	
	

	RETURN

END
GO

GRANT EXECUTE ON dbo.ITA_SP_UPDATE_DATISPEDIZIONE TO Metodo98
GO
