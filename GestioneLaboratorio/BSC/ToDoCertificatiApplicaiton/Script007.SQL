
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
	
	, dateadd(N, (6 * 60) + (rn * 30), @DATAPOSIZIONAMENTO) AS Start
	, dateadd(N, (6 * 60) + ((rn + 1) * 30), @DATAPOSIZIONAMENTO) AS [End] 
	FROM (
	SELECT 
	ROW_NUMBER() OVER(PARTITION BY e.BackgroundId ORDER BY RAGIONESOCIALE) AS RN
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
, @giorno SMALLINT = 0
)
AS

BEGIN

	SET NOCOUNT ON
	
	IF @idtesta > 0
	BEGIN
		-- aggiornamento dati extra
		UPDATE EXTRARIGHEDOC
		SET DATAPOSIZIONAMENTO = convert(DATE, @dataposizionamento)
		WHERE IDTESTA = @idtesta AND idriga = @idriga
		
		DELETE FROM Appointments 
		WHERE 
			mkey = CAST(@idtesta AS VARCHAR) + '|' + CAST(@idriga AS VARCHAR)
	END 	
	
	IF @giorno = 1
	BEGIN
		-- appuntamenti
		DELETE FROM Appointments 
		WHERE 
			mkey = CAST(@idtesta AS VARCHAR) + '|' + CAST(@idriga AS VARCHAR)
			OR convert(DATE, start) = convert(DATE, @dataposizionamento)
			
		
		INSERT INTO dbo.Appointments (Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey, TIPOIMBALLO, TIPOORDINE)
		SELECT 
		Summary, Start, [End], RecurrenceRule, MasterEventId, Location, Description, BackgroundId, mkey, TIPOIMBALLO, TIPOORDINE 
		FROM GET_PORTAFOGLIOORDINI_APPOINTMENTS(convert(DATE, @dataposizionamento)) -- idtesta = @idtesta AND idriga = @idriga	
	END	
	

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
	, DATAPOSIZIONAMENTO = convert(DATE, @dataposizionamento)
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
	EXEC ITA_SP_UPDATE_DATEPOSIZIONAMENTO @idtesta, @idriga, @dataposizionamento, 1
	
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







IF OBJECT_ID ('dbo.GET_DISPONIBILITAMETODO') IS NOT NULL
	DROP FUNCTION dbo.GET_DISPONIBILITAMETODO
GO

/*
	Autore:	Mantovani Daniele
	Data:	03/02/2004
	Descr.:	Ritorna l'elenco dei valori di un certo attributo (di tipo elenco) e se ciascun valore dipende da un certo valore parent.
		
	Modifiche:

*/

CREATE FUNCTION [dbo].[GET_DISPONIBILITAMETODO]
    (
      @codart VARCHAR(50)
    )
RETURNS TABLE
AS RETURN
    ( 
		SELECT CODDEPOSITO as Deposito, Max(datamov) AS DataMov, SUM(Carico) AS Carichi,SUM(ResoDaScarico) AS ResiDaScarico,SUM(Scarico) AS Scarichi,SUM(ResoDaCarico) AS ResiDaCarico,SUM(Ordinato) AS Ordinato,SUM(Impegnato) AS Impegnato 
		FROM VistaGiacenze 
		WHERE CodArt=@codart AND CodDeposito IN (select CODICE from ANAGRAFICADEPOSITI where DISPONIBILE=1) 
		AND DataMov >= CAST(year(getdate()) AS VARCHAR) + '0101' AND DataMov < CAST(year(getdate()) + 1 AS VARCHAR) + '0101' AND Esercizio=2017 
		GROUP BY CodArt,CodDeposito

    )




GO


GRANT ALL ON GET_DISPONIBILITAMETODO TO metodo98
GO



IF OBJECT_ID ('dbo.ZSI_VISTA_GIACENZEBERTELLO_DETTAGLIO') IS NOT NULL
	DROP VIEW dbo.ZSI_VISTA_GIACENZEBERTELLO_DETTAGLIO
GO

create view ZSI_VISTA_GIACENZEBERTELLO_DETTAGLIO as 
SELECT
 t3.codice
 , t3.descr
 , t2.nome as lotto 
 , CAST((t1.giacenza * t1.peso_unitario) AS DECIMAL(16,6)) AS GIACENZA
 , t1.data_alloc as datamovimento
 , t5.descr as imballaggio
 , t1.peso_unitario
--, t3.* 
--, t2.* 
--, t1.* 
  FROM [SRV2K3MA].[ZSI].[dbo].[Mappa] t1 join [SRV2K3MA].[ZSI].[dbo].[Lotto] t2 on t2.[id] = t1.[id_lotto] 
join [SRV2K3MA].[ZSI].[dbo].[Articolo] t3 on t3.[id] = t2.id_articolo 
join [SRV2K3MA].[ZSI].[dbo].[Confezione] t4 on t4.id_mappa = t1.id
join [SRV2K3MA].[ZSI].[dbo].[T_Confezione] t5 on t5.id = t4.id_tipo_conf
WHERE t1.qta_prenotata_prel = 0
GO

GRANT SELECT ON dbo.ZSI_VISTA_GIACENZEBERTELLO_DETTAGLIO TO Metodo98
GO

