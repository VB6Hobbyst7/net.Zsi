



ALTER TABLE Appointments ADD TipoImballo VARCHAR(5)
GO
ALTER TABLE Appointments ADD TipoOrdine VARCHAR(5)
GO


IF OBJECT_ID ('dbo.EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS') IS NOT NULL
	DROP VIEW dbo.EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS
GO

create view EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS as 


SELECT 
e.DESCRIZIONE as Summary
--, rn
--, dateadd(N, (8 * 60) + (rn * 30), DATAPOSIZIONAMENTO) AS Start
--, dateadd(N, (8 * 60) + ((rn + 1) * 30), DATAPOSIZIONAMENTO) AS [End]
, '' AS RecurrenceRule
, 0 as MasterEventId
, e.DESTINAZIONE AS Location
, e.DESCRIZIONE 
	+ ' Cliente: ' + e.RAGIONESOCIALE 
	+ ' Q.tà: ' + CAST(CAST(XLS_QTA AS INT) AS VARCHAR) 
	+ ' [' + IMBALLO + '] Lotto: ' + NRLOTTO 
	+ ' Ordine: ' + CAST(e.NUMERODOC AS VARCHAR)
	+ ' Fusti/Cubi: ' + CAST(e.NRPEZZIIMBALLO AS VARCHAR)
	
	AS Description
--, (CASE WHEN e.MAGSTATORIGA = 1 THEN 4 
--WHEN e.MAGSTATORIGA = 2 THEN 11
--WHEN e.MAGSTATORIGA = 3 THEN 2
--ELSE 1 END) AS BackgroundId
, (CASE WHEN E.TIPODOC = 'OCE' THEN 7 ELSE 1 END) AS BackgroundId
, CAST(e.idtesta AS VARCHAR) + '|' + CAST(e.idriga AS VARCHAR) AS MKEY
, e.IDTESTA, e.idriga
, e.DATACARICO
, e.DATAPOSIZIONAMENTO
, (CASE WHEN IMBALLO = 'CISTERNA' THEN 'C' ELSE 'I' END) AS TIPOIMBALLO
, (CASE WHEN E.TIPODOC = 'OCE' THEN 'E' ELSE 'I' END) AS TIPOORDINE
, E.RAGIONESOCIALE
FROM EXCEL_PORTAFOGLIOORDINI e
GO

GRANT DELETE ON dbo.EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO
GRANT INSERT ON dbo.EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO
GRANT REFERENCES ON dbo.EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO
GRANT SELECT ON dbo.EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO
GRANT UPDATE ON dbo.EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO




IF OBJECT_ID ('dbo.GET_PORTAFOGLIOORDINI_APPOINTMENTS') IS NOT NULL
	DROP FUNCTION dbo.GET_PORTAFOGLIOORDINI_APPOINTMENTS
GO

CREATE FUNCTION dbo.GET_PORTAFOGLIOORDINI_APPOINTMENTS
(
@DATAPOSIZIONAMENTO DATETIME
)
RETURNS TABLE AS RETURN
(

	SELECT *
	
	, dateadd(N, (7 * 60) + (rn * 30), @DATAPOSIZIONAMENTO) AS Start
	, dateadd(N, (7 * 60) + ((rn + 1) * 30), @DATAPOSIZIONAMENTO) AS [End] 
	FROM (
	SELECT 
	ROW_NUMBER() OVER(ORDER BY RAGIONESOCIALE) AS RN
	, * FROM EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS e
	
	WHERE 	convert(DATE, E.dataposizionamento) = @DATAPOSIZIONAMENTO
	) X
)
GO

GRANT DELETE ON dbo.GET_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO
GRANT INSERT ON dbo.GET_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO
GRANT REFERENCES ON dbo.GET_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO
GRANT SELECT ON dbo.GET_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO
GRANT UPDATE ON dbo.GET_PORTAFOGLIOORDINI_APPOINTMENTS TO Metodo98
GO



IF OBJECT_ID ('dbo.ITA_SP_UPDATE_DATEPOSIZIONAMENTO') IS NOT NULL
	DROP PROCEDURE dbo.ITA_SP_UPDATE_DATEPOSIZIONAMENTO
GO

CREATE PROCEDURE [dbo].[ITA_SP_UPDATE_DATEPOSIZIONAMENTO]
(

@idtesta DECIMAL(10)
, @idriga DECIMAL(5)
, @dataposizionamento DATETIME
)
AS

BEGIN

	SET NOCOUNT ON
	
	-- aggiornamento dati extra
	UPDATE EXTRARIGHEDOC
	SET DATAPOSIZIONAMENTO = @dataposizionamento
	WHERE IDTESTA = @idtesta AND idriga = @idriga
	
	
	
	-- appuntamenti
	DELETE FROM Appointments 
	WHERE 
		mkey = CAST(@idtesta AS VARCHAR) + '|' + CAST(@idriga AS VARCHAR)
		OR convert(DATE, start) = @dataposizionamento
		
	
	INSERT INTO dbo.Appointments (Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey, TIPOIMBALLO, TIPOORDINE)
	SELECT 
	Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey, TIPOIMBALLO, TIPOORDINE 
	FROM GET_PORTAFOGLIOORDINI_APPOINTMENTS(@dataposizionamento) -- idtesta = @idtesta AND idriga = @idriga	
	
	

	RETURN

END
GO



GRANT EXECUTE ON dbo.ITA_SP_UPDATE_DATEPOSIZIONAMENTO TO Metodo98
GO










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
, @user VARCHAR(255) = ''
, @machine VARCHAR(255) = ''

)
AS

BEGIN

	SET NOCOUNT ON
	
	UPDATE  ITA_ARCHIVIO_DOCUMENTI 
	SET TestPubblicazione = 1 
	WHERE Progressivo = @idtesta
	

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
	DELETE FROM Appointments 
	WHERE 
		mkey = CAST(@idtesta AS VARCHAR) + '|' + CAST(@idriga AS VARCHAR)
		OR convert(DATE, start) = @dataposizionamento
		
	
--	INSERT INTO dbo.Appointments (Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey, TIPOIMBALLO, TIPOORDINE)
--	SELECT 
--	Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey, TIPOIMBALLO, TIPOORDINE 
--	FROM EXCEL_PORTAFOGLIOORDINI_APPOINTMENTS e 
--	WHERE convert(DATE, e.dataposizionamento) = @dataposizionamento -- idtesta = @idtesta AND idriga = @idriga	

	INSERT INTO dbo.Appointments (Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey, TIPOIMBALLO, TIPOORDINE)
	SELECT 
	Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey, TIPOIMBALLO, TIPOORDINE 
	FROM GET_PORTAFOGLIOORDINI_APPOINTMENTS(@dataposizionamento) -- idtesta = @idtesta AND idriga = @idriga		
	
	-- LOG
	INSERT INTO LOG_EXTRARIGHEDOCM05
	(
		IDTESTA,
		IDRIGA,
		UTENTEMODIFICA,
		DATAMODIFICA,
		UTENTE,
		MACHINE,
		ACTION
	)
	VALUES 
	(
		@idtesta,
		@idriga,
		'',
		GETDATE(),
		@user,
		@machine,
		'AGGIORNAMENTO DA M05'
	)
	
	
	RETURN

END
GO

GRANT EXECUTE ON dbo.ITA_SP_UPDATE_DATISPEDIZIONE TO Metodo98
GO


IF OBJECT_ID ('dbo.TU_Appointmets') IS NOT NULL
	DROP TRIGGER dbo.TU_Appointmets
GO

/*  INSERT TRIGGER "TU_Appointmets" FOR TABLE "UBICAZIONIARTICOLI"  */
CREATE TRIGGER TU_Appointmets ON Appointments FOR UPDATE AS
BEGIN
    DECLARE
       @MAXCARD  INT,
       @NUMROWS  INT,
       @NUMNULL  INT,
       @ERRNO    INT,
       @ERRMSG   VARCHAR(255)
    SELECT  @NUMROWS = @@ROWCOUNT
    IF @NUMROWS = 0
       RETURN
    
    IF UPDATE(START)
    BEGIN
		UPDATE E
		SET E.DATAPOSIZIONAMENTO = CONVERT(DATE, A.START)
		--SELECT * 
		FROM INSERTED a, extrarighedoc e
		WHERE a.MKEY = CAST(e.IDTESTA AS VARCHAR) + '|' + CAST(e.IDRIGA AS VARCHAR) 
		AND CONVERT(DATE, a.START) <> E.DATAPOSIZIONAMENTO
    END
    
    RETURN
/*  ERRORS HANDLING  */
ERROR:
    RAISERROR (@ERRMSG, 1, 1)
    ROLLBACK  TRANSACTION
END
GO








/*

	DECLARE @idtesta DECIMAL(10)
	DECLARE @idriga DECIMAL(5)
	DECLARE @dataposizionamento datetime
	
	DECLARE rSqlA CURSOR LOCAL KEYSET FOR 
	SELECT idtesta, idriga, dataposizionamento
	FROM EXCEL_PORTAFOGLIOORDINI
	
	
	
	
	OPEN rSqlA

	FETCH NEXT from rSqlA INTO @idtesta, @idriga, @dataposizionamento
	WHILE (@@FETCH_STATUS <> -1)
		BEGIN
		
			EXECUTE dbo.ITA_SP_UPDATE_DATEPOSIZIONAMENTO @idtesta, @idriga, @dataposizionamento
			
			FETCH NEXT from rSqlA INTO @idtesta, @idriga, @dataposizionamento
		END

	CLOSE rSqlA
	DEALLOCATE rSqlA
	
	*/