

IF OBJECT_ID ('dbo.fn_GeTCRMCOLUMNS') IS NOT NULL
	DROP FUNCTION dbo.fn_GeTCRMCOLUMNS
GO

create function [dbo].[fn_GeTCRMCOLUMNS](@V AS VARCHAR(255)) returns varchar(8000) as begin
	declare @Str varchar(8000)
	
	set @Str = ''
	
	select @Str = @Str + (case @Str when '' then '' else ',' end) + I.COLUMN_NAME --+ ''''  + I.COLUMN_NAME + ''' AS ' + I.COLUMN_NAME
	FROM INFORMATION_SCHEMA.COLUMNS I
	WHERE TABLE_NAME = @V
	--ORDER BY I.ORDINAL_POSITION
	group by ORDINAL_POSITION, COLUMN_NAME

	IF @str = ''
	BEGIN
		-- imposto a 0 altrimenti la pubblicazione va in errore 
		SET @str = '0'
	END

	return @Str
end
GO

GRANT EXECUTE ON dbo.fn_GeTCRMCOLUMNS TO Metodo98
GO

GRANT REFERENCES ON dbo.fn_GeTCRMCOLUMNS TO Metodo98
GO


IF OBJECT_ID ('dbo.EXCEL_EXPORT2CRM_AGENTI') IS NOT NULL
	DROP VIEW dbo.EXCEL_EXPORT2CRM_AGENTI
GO

CREATE  VIEW EXCEL_EXPORT2CRM_AGENTI AS
SELECT AA.CODagente, AA.DscAgente, AA.PartitaIVA, AA.CodFiscale, AA.TELEX EMAIL, AE.PASSWORD--, 9 AS I  
FROM 
Anagraficaagenti AA LEFT OUTER JOIN AGENTI_ECOM AE ON AA.CODagente = AE.CODagente
GO

GRANT SELECT ON dbo.EXCEL_EXPORT2CRM_AGENTI TO Metodo98
GO


IF OBJECT_ID ('dbo.EXCEL_EXPORT2CRM_ARTICOLI') IS NOT NULL
	DROP VIEW dbo.EXCEL_EXPORT2CRM_ARTICOLI
GO

create view EXCEL_EXPORT2CRM_ARTICOLI as 
/*
SELECT 
	'CODICE' AS CODICE,
	'DESCRIZIONE' AS DESCRIZIONE,
	'NRPEZZIIMBALLO' AS NRPEZZIIMBALLO,
	'NOMENCLCOMBINATA1' AS NOMENCLCOMBINATA1,
	'PROVENIENZA' AS PROVENIENZA,
	'GRUPPO' AS GRUPPO,
	'NATURA' AS NATURA,
	'CATEGORIASTAT' AS CATEGORIASTAT,
	'FAMIGLIA' AS FAMIGLIA,
	'CONCENTRAZIONE' AS CONCENTRAZIONE,
	'SPECIALITIES' AS SPECIALITIES,
	'ECOCERT' AS ECOCERT,
	'COSMOS' AS COSMOS,
	0 AS I
UNION
*/
SELECT 
	AA.CODICE
	, AA.DESCRIZIONE
	, AA.NRPEZZIIMBALLO
--	, (SELECT TOP 1 I.DESCRIZIONE FROM TABIMBALLI I WHERE I.CODICE = AA.RIFERIMIMBALLO) AS RIFERIMENTOIMBALLO
	, AA.NOMENCLCOMBINATA1
	, (CASE WHEN AP.PROVENIENZA = 2 THEN 'CONTO LAVORO'
		WHEN AP.PROVENIENZA = 1 THEN 'PRODUZIONE'
		ELSE 'ACQUISTO'
		END) AS PROVENIENZA
	, EMC.GRUPPO
	, EMC.NATURA
	, EMC.CATEGORIASTAT
	, EMC.FAMIGLIA
	, EMC.CONCENTRAZIONE
	, EMC.SPECIALITIES
	, EMC.ECOCERT
	, EMC.COSMOS
	--, 9 AS I
FROM ANAGRAFICAARTICOLI AA WITH (NOLOCK) JOIN ANAGRAFICAARTICOLIPROD AP WITH (NOLOCK) ON AP.CODICEART = AA.CODICE 
JOIN EXTRAMAGCRM EMC WITH (NOLOCK) ON EMC.CODART = AA.CODICE
WHERE AP.ESERCIZIO = YEAR(GETDATE())
GO



IF OBJECT_ID ('dbo.EXCEL_EXPORT2CRM_CONTATTICLIENTI') IS NOT NULL
	DROP VIEW dbo.EXCEL_EXPORT2CRM_CONTATTICLIENTI
GO

create view EXCEL_EXPORT2CRM_CONTATTICLIENTI as 

SELECT 
	--a.TIPOCONTO,
	t.RIFCODCONTO AS CODICEAZIENDA
	, T.CODICE
	, T.COGNOME
	, T.TELEFONO
	, T.CELL
	, T.EMAIL
	, tr.DESCRIZIONE AS RUOLOCONTATTO
  FROM ANAGRAFICACF a 
LEFT OUTER JOIN TABELLAPERSONALE t with (nolock) ON a.CODCONTO = t.RIFCODCONTO
LEFT OUTER JOIN TabellaRuoli Tr with (nolock) ON TR.CODICE = t.CODRUOLO

WHERE a.TIPOCONTO = 'c' AND NOT t.RIFCODCONTO IS NULL
GO

GRANT SELECT ON dbo.EXCEL_EXPORT2CRM_CONTATTICLIENTI TO Metodo98
GO






IF OBJECT_ID ('dbo.EXCEL_EXPORT2CRM_DOCUMENTI_ORDINI_TESTE') IS NOT NULL
	DROP VIEW dbo.EXCEL_EXPORT2CRM_DOCUMENTI_ORDINI_TESTE
GO

create view EXCEL_EXPORT2CRM_DOCUMENTI_ORDINI_TESTE as 

SELECT DISTINCT
	CODCLIFOR,
	RAGIONESOCIALE,
	CODAGENTE1,
	AGENTE,
	DOCUMENTO,
	TIPODOCUMENTO,
	NUMRIFDOC,
	DATARIFDOC,
	ESERCIZIO,
	TIPODOC,
	NUMERODOC,
	BIS,
	DATADOC,
	INCOTERM,
	DESCRIZIONEPAGAMENTO,
	DATASCADENZA,
	IMPORTOSCEURO,
	DESTINAZIONE,
	IDOBJECT,
	IDTESTA
	EXPORTCRM
FROM dbo.EXCEL_EXPORT2CRM_DOCUMENTI_ORDINI
GO

GRANT SELECT ON dbo.EXCEL_EXPORT2CRM_DOCUMENTI_ORDINI_TESTE TO Metodo98
GO


IF OBJECT_ID ('dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE') IS NOT NULL
	DROP VIEW dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE
GO

create view EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE as 
SELECT 
	T.CODCLIFOR
		, AC.DSCCONTO1 AS RAGIONESOCIALE
		, T.CODAGENTE1
		, (SELECT TOP 1 AA.DSCAGENTE FROM ANAGRAFICAAGENTI AA WHERE AA.CODAGENTE = T.CODAGENTE1) AS AGENTE
        , T.TIPODOC + '/' + CAST(T.ESERCIZIO AS VARCHAR) + '/' + CAST(T.NUMERODOC AS VARCHAR) AS DOCUMENTO -- + '/' + CAST(R.POSIZIONE AS VARCHAR) AS DOCUMENTO 
        , P.DESCRIZIONE AS TIPODOCUMENTO
        , T.NUMRIFDOC
        , T.DATARIFDOC
        , R.CODART AS ARTICOLO
        , R.DESCRIZIONEART AS DESCRIZIONE
 		, I.DESCRIZIONE IMBALLO 
        , CAST(R.QTAGEST AS INT) AS QUANTITA
        , R.UMGEST
        , R.PREZZOUNITLORDOEURO
        , R.SCONTIESTESI
        , (SELECT TOP 1 TRATTAMENTOIVA.DESCRIZIONE FROM TRATTAMENTOIVA WHERE TRATTAMENTOIVA.CODICE =   R.CODIVA) AS IVA
        , R.PREZZOUNITNETTOEURO
        , R.DATACONSEGNA
        , R.TOTLORDORIGAEURO
        , R.TOTNETTORIGAEURO
        , T.ESERCIZIO, T.TIPODOC, T.NUMERODOC, T.BIS, T.DATADOC 
		, R.ANNOTAZIONI
		, (SELECT TOP 1 TABPORTO.DESCRIZIONE FROM TABPORTO WHERE TABPORTO.CODICE = T.PORTO) AS INCOTERM
		, (SELECT TOP 1 TABPAGAMENTI.DESCRIZIONE FROM TABPAGAMENTI WHERE TABPAGAMENTI.CODICE = t.codpagamento) AS DESCRIZIONEPAGAMENTO
		, VS.DATASCADENZA, VS.IMPORTOSCEURO 
        , (CASE WHEN T.RAGSOCDDM IS NULL THEN 
        		AC.DSCCONTO1 + ' ' + AC.INDIRIZZO + ' ' + AC.LOCALITA + ' ' + AC.PROVINCIA 
        	ELSE
        		t.RAGSOCDDM + ' ' + t.INDIRIZZODDM + ' ' + t.LOCALITADDM + ' ' + t.PROVINCIADDM 
        	END) AS DESTINAZIONE
        , K.IdPubblicazioneKnos AS IDOBJECT
        , T.PROGRESSIVO AS IDTESTA, R.IDRIGA 
        , (SELECT TOP 1 1 FROM EXTRAMAGCRM X WHERE X.CODART = R.CODART) AS EXPORTCRM
        
FROM TESTEDOCUMENTI T WITH (NOLOCK) JOIN EXTRATESTEDOC ET WITH (NOLOCK) ON T.PROGRESSIVO = ET.IDTESTA 
JOIN PARAMETRIDOC P ON P.CODICE = T.TIPODOC
JOIN RIGHEDOCUMENTI R WITH (NOLOCK) ON T.PROGRESSIVO = R.IDTESTA 
JOIN ANAGRAFICACF AC WITH (NOLOCK) ON AC.CODCONTO = T.CODCLIFOR 
--JOIN EXTRACLIENTI WITH (NOLOCK) ON EXTRACLIENTI.CODCONTO = AC.CODCONTO
--JOIN ANAGRAFICAARTICOLI AA WITH (NOLOCK) ON AA.CODICE = R.CODART 
JOIN EXTRARIGHEDOC ER WITH (NOLOCK) ON ER.IDTESTA = R.IDTESTA AND ER.IDRIGA = R.IDRIGA 
LEFT OUTER JOIN VISTASCADENZE VS ON VS.TIPODOC = T.TIPODOC AND VS.ESERCIZIO = T.ESERCIZIO AND VS.NUMDOC = T.NUMERODOC AND VS.BIS = T.BIS
--LEFT OUTER JOIN ZSI_CONFEZSPEDIBILE z WITH (NOLOCK) ON Z.PROGRESSIVO = r.IDTESTA AND z.idriga = r.IDRIGA
LEFT OUTER JOIN TABIMBALLI I WITH (NOLOCK) ON R.CODIMBALLO = I.CODICE 
LEFT OUTER JOIN ITA_ARCHIVIO_DOCUMENTI K ON K.Progressivo = T.PROGRESSIVO
WHERE 
	--t.tipodoc IN ('OCC', 'OCE', 'OCG', 'FTI', 'FTE') AND 
	P.CLIFOR = 'c' AND P.TIPO IN ('F', 'N') AND
	R.CODART <> '' AND
	R.QTAGESTRES > 0 
	AND t.esercizio >= 2016
	AND isnull(VS.NUMSCAD, 1) = 1
GO

GRANT SELECT ON dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE TO Metodo98
GO



IF OBJECT_ID ('dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_TESTE') IS NOT NULL
	DROP VIEW dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_TESTE
GO

create view EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_TESTE as 

SELECT DISTINCT
	CODCLIFOR,
	RAGIONESOCIALE,
	CODAGENTE1,
	AGENTE,
	DOCUMENTO,
	TIPODOCUMENTO,
	NUMRIFDOC,
	DATARIFDOC,
	ESERCIZIO,
	TIPODOC,
	NUMERODOC,
	BIS,
	DATADOC,
	INCOTERM,
	DESCRIZIONEPAGAMENTO,
	DATASCADENZA,
	IMPORTOSCEURO,
	DESTINAZIONE,
	IDOBJECT,
	IDTESTA
	EXPORTCRM
FROM dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE
GO

GRANT SELECT ON dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_TESTE TO Metodo98
GO




IF OBJECT_ID ('dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_SCADENZE') IS NOT NULL
	DROP VIEW dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_SCADENZE
GO

create view EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_SCADENZE as 
SELECT 
	T.CODCLIFOR
        , T.TIPODOC + '/' + CAST(T.ESERCIZIO AS VARCHAR) + '/' + CAST(T.NUMERODOC AS VARCHAR) AS DOCUMENTO -- + '/' + CAST(R.POSIZIONE AS VARCHAR) AS DOCUMENTO 
        , P.DESCRIZIONE AS TIPODOCUMENTO
        , T.ESERCIZIO, T.TIPODOC, T.NUMERODOC, T.BIS, T.DATADOC 
		, VS.DATASCADENZA, VS.IMPORTOSCEURO 
		, vs.ESITO
		, (CASE WHEN vs.ESITO = 0 THEN 'NON EMESSO'
			WHEN VS.ESITO = 1 THEN 'EMESSO'
			WHEN VS.ESITO = 2 THEN 'PAGATO'
			WHEN VS.ESITO = 3 THEN 'INSOLUTO'
			WHEN VS.ESITO = 4 THEN 'INSOLUTO PAGATO'
			ELSE '' END) AS DSCESITO
		, VS.DATAPAGEFF
FROM TESTEDOCUMENTI T WITH (NOLOCK) 
JOIN PARAMETRIDOC P WITH (NOLOCK) ON P.CODICE = T.TIPODOC
INNER JOIN VISTASCADENZE VS WITH (NOLOCK) ON VS.TIPODOC = T.TIPODOC AND VS.ESERCIZIO = T.ESERCIZIO AND VS.NUMDOC = T.NUMERODOC AND VS.BIS = T.BIS
WHERE 
	--t.tipodoc IN ('OCC', 'OCE', 'OCG', 'FTI', 'FTE') AND 
	P.CLIFOR = 'c' AND P.TIPO IN ('F', 'N') 
	AND t.esercizio >= 2016
GO

GRANT SELECT ON dbo.EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_SCADENZE TO Metodo98
GO



SELECT * FROM EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_SCADENZE


USE [msdb]
GO

/****** Object:  Job [EXPORT2CRM]    Script Date: 02/28/2017 15:59:17 ******/
IF  EXISTS (SELECT job_id FROM msdb.dbo.sysjobs_view WHERE name = N'EXPORT2CRM')
EXEC msdb.dbo.sp_delete_job @job_id=N'2d13429d-2e2f-461c-ac7b-813a676c3b65', @delete_unused_schedule=1
GO

USE [msdb]
GO

/****** Object:  Job [EXPORT2CRM]    Script Date: 02/28/2017 15:59:17 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** Object:  JobCategory [[Uncategorized (Local)]]]    Script Date: 02/28/2017 15:59:17 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'EXPORT2CRM', 
		@enabled=1, 
		@notify_level_eventlog=0, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'Nessuna descrizione disponibile.', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [EXPORT_CLIENTI]    Script Date: 02/28/2017 15:59:18 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'EXPORT_CLIENTI', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=3, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'DECLARE @procedura AS VARCHAR(2000)
SET NOCOUNT ON
SET @procedura = ''SQLCMD -S . -d ZSI -W -h -1 -Q "SET NOCOUNT ON;SELECT DBO.fn_GeTCRMCOLUMNS(''''EXCEL_EXPORT2CRM_CLIENTI'''');SET NOCOUNT ON;SELECT * FROM EXCEL_EXPORT2CRM_CLIENTI" -s ";" -o "E:\MET\ITALCOM\CRM\EXPORT\clienti.csv"''
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''START'', GETDATE())
EXEC xp_cmdshell @procedura
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''STOP'', GETDATE())', 
		@database_name=N'ZSI', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [EXPORT_CONTATTICLIENTI]    Script Date: 02/28/2017 15:59:18 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'EXPORT_CONTATTICLIENTI', 
		@step_id=2, 
		@cmdexec_success_code=0, 
		@on_success_action=3, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'DECLARE @procedura AS VARCHAR(2000)
SET NOCOUNT ON
SET @procedura = ''SQLCMD -S . -d ZSI -W -h -1 -Q "SET NOCOUNT ON;SELECT DBO.fn_GeTCRMCOLUMNS(''''EXCEL_EXPORT2CRM_CONTATTICLIENTI'''');SET NOCOUNT ON;SELECT * FROM EXCEL_EXPORT2CRM_CONTATTICLIENTI" -s ";" -o "E:\MET\ITALCOM\CRM\EXPORT\contatticlienti.csv"''
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''START'', GETDATE())
EXEC xp_cmdshell @procedura
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''STOP'', GETDATE())', 
		@database_name=N'ZSI', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [EXPORT_AGENTI]    Script Date: 02/28/2017 15:59:18 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'EXPORT_AGENTI', 
		@step_id=3, 
		@cmdexec_success_code=0, 
		@on_success_action=3, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'DECLARE @procedura AS VARCHAR(2000)
SET NOCOUNT ON
SET @procedura = ''SQLCMD -S . -d ZSI -W -h -1 -Q "SET NOCOUNT ON;SELECT DBO.fn_GeTCRMCOLUMNS(''''EXCEL_EXPORT2CRM_AGENTI'''');SET NOCOUNT ON;SELECT * FROM EXCEL_EXPORT2CRM_AGENTI" -s ";" -o "E:\MET\ITALCOM\CRM\EXPORT\agenti.csv"''
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''START'', GETDATE())
EXEC xp_cmdshell @procedura
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''STOP'', GETDATE())', 
		@database_name=N'ZSI', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [EXPORT_ARTICOLI]    Script Date: 02/28/2017 15:59:18 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'EXPORT_ARTICOLI', 
		@step_id=4, 
		@cmdexec_success_code=0, 
		@on_success_action=3, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'SET NOCOUNT ON;
DECLARE @procedura AS VARCHAR(2000)
SET NOCOUNT ON
SET @procedura = ''SQLCMD -S . -d ZSI -W -h -1 -Q "SET NOCOUNT ON;SELECT DBO.fn_GeTCRMCOLUMNS(''''EXCEL_EXPORT2CRM_ARTICOLI'''');SET NOCOUNT ON;SELECT * FROM EXCEL_EXPORT2CRM_ARTICOLI ORDER BY I" -s ";" -o "E:\MET\ITALCOM\CRM\EXPORT\articoli.csv"''
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''START'', GETDATE())
EXEC xp_cmdshell @procedura
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''STOP'', GETDATE())
', 
		@database_name=N'ZSI', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [EXPORT_ORDINI]    Script Date: 02/28/2017 15:59:18 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'EXPORT_ORDINI', 
		@step_id=5, 
		@cmdexec_success_code=0, 
		@on_success_action=3, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'
DECLARE @procedura AS VARCHAR(2000)
SET NOCOUNT ON
SET @procedura = ''SQLCMD -S . -d ZSI -W -h -1 -Q "SET NOCOUNT ON;SELECT DBO.fn_GeTCRMCOLUMNS(''''EXCEL_EXPORT2CRM_DOCUMENTI_ORDINI_TESTE'''');SET NOCOUNT ON;SELECT DISTINCT * FROM EXCEL_EXPORT2CRM_DOCUMENTI_ORDINI_TESTE" -s ";" -o "E:\MET\ITALCOM\CRM\EXPORT\ordini_teste.csv"''
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''START'', GETDATE())
EXEC xp_cmdshell @procedura
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''STOP'', GETDATE())




SET NOCOUNT ON
SET @procedura = ''SQLCMD -S . -d ZSI -W -h -1 -Q "SET NOCOUNT ON;SELECT DBO.fn_GeTCRMCOLUMNS(''''EXCEL_EXPORT2CRM_DOCUMENTI_ORDINI'''');SET NOCOUNT ON;SELECT * FROM EXCEL_EXPORT2CRM_DOCUMENTI_ORDINI" -s ";" -o "E:\MET\ITALCOM\CRM\EXPORT\ordini_righe.csv"''
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''START'', GETDATE())
EXEC xp_cmdshell @procedura
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''STOP'', GETDATE())
', 
		@database_name=N'ZSI', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [EXPORT_FATTURE]    Script Date: 02/28/2017 15:59:18 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'EXPORT_FATTURE', 
		@step_id=6, 
		@cmdexec_success_code=0, 
		@on_success_action=3, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'
DECLARE @procedura AS VARCHAR(2000)
SET NOCOUNT ON
SET @procedura = ''SQLCMD -S . -d ZSI -W -h -1 -Q "SET NOCOUNT ON;SELECT DBO.fn_GeTCRMCOLUMNS(''''EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_TESTE'''');SET NOCOUNT ON;SELECT DISTINCT * FROM EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE_TESTE" -s ";" -o "E:\MET\ITALCOM\CRM\EXPORT\fatture_teste.csv"''
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''START'', GETDATE())
EXEC xp_cmdshell @procedura
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''STOP'', GETDATE())




SET NOCOUNT ON
SET @procedura = ''SQLCMD -S . -d ZSI -W -h -1 -Q "SET NOCOUNT ON;SELECT DBO.fn_GeTCRMCOLUMNS(''''EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE'''');SET NOCOUNT ON;SELECT * FROM EXCEL_EXPORT2CRM_DOCUMENTI_FATTURE" -s ";" -o "E:\MET\ITALCOM\CRM\EXPORT\fatture_righe.csv"''
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''START'', GETDATE())
EXEC xp_cmdshell @procedura
INSERT INTO dbo.EXPORT2CRM_LOG (PROCEDURA, ESITO, datamodifica) VALUES (@procedura,  ''STOP'', GETDATE())
', 
		@database_name=N'ZSI', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [CLEAN_START]    Script Date: 02/28/2017 15:59:18 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'CLEAN_START', 
		@step_id=7, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'DELETE FROM EXPORT2CRM_LOG WHERE ESITO = ''START''', 
		@database_name=N'ZSI', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'EXPORT2CRM', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20170208, 
		@active_end_date=99991231, 
		@active_start_time=10000, 
		@active_end_time=235959, 
		@schedule_uid=N'edae8cb0-f954-476e-a3ad-53eba2adccef'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO



