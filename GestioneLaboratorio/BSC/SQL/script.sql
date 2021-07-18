
/* KNOS


if not(object_id(N'Metodo_View_LinkageBollettini') is null)
DROP VIEW [dbo].[Metodo_View_LinkageBollettini]
GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[Metodo_View_LinkageBollettini]
AS

select 
	ol.idparent AS IDOBJECT_BOL 
	, ol.IdChild AS IDOBJECT_ART
	, olcli.IdChild AS IDOBJECT_CLIFOR
	, od.*
from 
	object_linkage ol 
	inner join Object_Linkage olcli on ol.IdParent = olcli.IdParent and ol.idchild <> olcli.IdChild
	inner join object_doc od on od.idobject = ol.IdParent
where ol.idattr = 5036 and olcli.IdAttr = 19
	AND OD.VERSION = OD.CURRENTVERSION
GO
*/



/*

IF OBJECT_ID ('dbo.fn_GetNOTIFICHEBSC') IS NOT NULL
	DROP FUNCTION dbo.fn_GetNOTIFICHEBSC
GO

CREATE    FUNCTION [dbo].[fn_GetNOTIFICHEBSC]() 
RETURNS @TEMP_NOTIFICHEBSC TABLE 
(
	CODCLIFOR           VARCHAR (7) NOT NULL,
	CODART              VARCHAR (50) NOT NULL,
	ARTICOLOBSC         VARCHAR (50) NOT NULL,
	CODLINGUA           VARCHAR (10) NOT NULL,
	IDOBJECT_CLI        INT NOT NULL,
	IDOBJECT_ART        INT NOT NULL,
	IDOBJECT_BOL        INT NOT NULL,
	IDOBJECT_SCH        INT NOT NULL,
	RAGIONESOCIALE      VARCHAR (255),
	DESCRIZIONEARTICOLO VARCHAR (255),
	EMAIL_CLIENTE       VARCHAR (1000),
	EMAIL_AGENTE        VARCHAR (1000),
	IDDOC               INT NOT NULL,
	FILENAME            VARCHAR (255),
	CODICEDOCUMENTO     VARCHAR (255),
	DATAULTIMOINVIO     DATETIME NULL,
	PRIMARY KEY (CODCLIFOR, CODART)
	WITH (FILLFACTOR = 90)
)
AS
BEGIN
	INSERT INTO @TEMP_NOTIFICHEBSC (CODCLIFOR, CODART, ARTICOLOBSC, CODLINGUA, IDOBJECT_CLI, IDOBJECT_ART, IDOBJECT_BOL, IDOBJECT_SCH, RAGIONESOCIALE, DESCRIZIONEARTICOLO, EMAIL_CLIENTE, EMAIL_AGENTE, IDDOC, FILENAME, CODICEDOCUMENTO, DATAULTIMOINVIO)
	SELECT CODCLIFOR, CODART, ARTICOLOBSC
		, '' AS CODLINGUA
		, 0 AS IDOBJECT_CLI
		, 0 AS IDOBJECT_ART
		, 0 AS IDOBJECT_BOL
		, 0 AS IDOBJECT_SCH
		, '' AS RAGIONESOCIALE
		, '' AS DESCRIZIONEARTICOLO
		, '' AS EMAIL_CLIENTE
		, '' AS EMAIL_AGENTE
		, 0 AS IDDOC
		, '' AS FILENAME
		, '' AS CODICEDOCUMENTO
		, NULL AS DATAULTIMOINVIO
	FROM VistaRelazioniCFVdaDocZSI_BSC
	
	UPDATE 
		@TEMP_NOTIFICHEBSC
	SET 
		IDOBJECT_CLI = isnull((SELECT TOP 1 i.IdPubblicazioneKnos FROM ITA_ARCHIVIO_CLIFOR i WHERE i.CodConto = CODCLIFOR), 0)

	UPDATE 
		@TEMP_NOTIFICHEBSC
	SET 
		IDOBJECT_ART = isnull((SELECT TOP 1 i.IdPubblicazioneKnos FROM ITA_ARCHIVIO_ARTICOLI i WHERE i.CODICE = ARTICOLOBSC), 0)

	UPDATE 
		@TEMP_NOTIFICHEBSC
	SET 
		DESCRIZIONEARTICOLO = (SELECT TOP 1 a.DESCRIZIONE FROM ANAGRAFICAARTICOLI A WHERE CODICE = ARTICOLOBSC)
	
	
	UPDATE  T 
	SET 
		T.RAGIONESOCIALE = a.DSCCONTO1,
		T.CODLINGUA = ISNULL(TN.DESCRIZIONELINGUA9, 'IT-')
	FROM 
		@TEMP_NOTIFICHEBSC T JOIN ANAGRAFICACF A ON A.CODCONTO = T.CODCLIFOR
		JOIN TABNAZIONI TN ON A.CODNAZIONE = TN.CODICE
		
	UPDATE T 
	SET 
		T.IDDOC = ISNULL(bol.IDDOC, 0)
		, T.FILENAME = bol.FILENAME
		, T.IDOBJECT_BOL = bol.IDOBJECT_BOL
	FROM
		@TEMP_NOTIFICHEBSC T JOIN Knos7_ZSI.DBO.Metodo_View_LinkageBollettini bol 
		ON bol.IDOBJECT_ART = T.IDOBJECT_ART and bol.IDOBJECT_CLIFOR = T.IDOBJECT_CLI

	UPDATE T 
	SET 
		T.IDDOC = ISNULL(bol.IDDOC, 0)
		, T.FILENAME = bol.FILENAME
		, T.IDOBJECT_BOL = bol.IDOBJECT_BOL
	FROM
		@TEMP_NOTIFICHEBSC T JOIN Knos7_ZSI.DBO.Metodo_View_LinkageBollettini bol 
		ON bol.IDOBJECT_ART = T.IDOBJECT_ART
	WHERE 
		T.IDDOC = 0


	UPDATE 
		@TEMP_NOTIFICHEBSC
	SET 
		EMAIL_CLIENTE = [dbo].[fn_GetEMAILCONTATTICLIENTE](CODCLIFOR)
	
	RETURN
END
GO


GRANT REFERENCES ON dbo.fn_GetNOTIFICHEBSC TO Metodo98
GO
GRANT SELECT ON dbo.fn_GetNOTIFICHEBSC TO Metodo98
GO



--SELECT * FROM dbo.fn_GetNOTIFICHEBSC()
*/



if not exists(select 1 from information_schema.columns where table_name = 'RELAZIONICFV' and column_name = 'ARTICOLOBSC')
	alter table RELAZIONICFV add ARTICOLOBSC VARCHAR(50)
go

IF OBJECT_ID ('dbo.VISTARELAZIONICFV') IS NOT NULL
	DROP VIEW dbo.VISTARELAZIONICFV
GO

CREATE VIEW [dbo].[VISTARELAZIONICFV]
AS
SELECT     CODCLIFOR, RIFERIMENTO, ARTICOLO, VARIANTI, DESCRIZIONE, TIPOREL, MOSTRAVARIANTI, UTENTEMODIFICA, DATAMODIFICA, 
NOTE,
CODBONUS,
DSCART_ALT,
                      (CASE WHEN VARIANTI = '' OR
                      VARIANTI IS NULL OR
                      CHARINDEX('?', VARIANTI) > 0 THEN
                          (SELECT     DESCRIZIONE
                            FROM          ANAGRAFICAARTICOLI
                            WHERE      CODICE = ARTICOLO) ELSE
                          (SELECT     DESCRIZIONE
                            FROM          ANAGRAFICAARTICOLI
                            WHERE      CODICE = ARTICOLO + '#' + VARIANTI) END) AS DSCART, (CASE WHEN VARIANTI = '' OR
                      VARIANTI IS NULL OR
                      CHARINDEX('?', VARIANTI) > 0 THEN
                          (SELECT     ARTTIPOLOGIA
                            FROM          ANAGRAFICAARTICOLI
                            WHERE      CODICE = ARTICOLO) ELSE
                          (SELECT     ARTTIPOLOGIA
                            FROM          ANAGRAFICAARTICOLI
                            WHERE      CODICE = ARTICOLO + '#' + VARIANTI) END) AS ARTTIPOLOGIA, (CASE WHEN VARIANTI = '' OR
                      VARIANTI IS NULL OR
                      CHARINDEX('?', VARIANTI) > 0 THEN
                          (SELECT     CODICE
                            FROM          ANAGRAFICAARTICOLI
                            WHERE      CODICE = ARTICOLO) ELSE
                          (SELECT     CODICEPRIMARIO
                            FROM          ANAGRAFICAARTICOLI
                            WHERE      CODICE = ARTICOLO + '#' + VARIANTI) END) AS CODICEPRIMARIO, (CASE WHEN TIPOREL = 1 THEN ARTICOLO ELSE RIFERIMENTO END) 
                      AS CODICERIF, (CASE WHEN Varianti = '' OR
                      VARIANTI IS NULL THEN Articolo ELSE (Articolo + '#' + replace(Varianti, '?', '%')) END) CodArticolo, escludiperperiodo
, RELAZIONICFV.ARTICOLOBSC
FROM         RELAZIONICFV


GO

GRANT SELECT ON dbo.VISTARELAZIONICFV TO Metodo98
GO


IF OBJECT_ID ('dbo.VistaRelazioniCFVdaDocZSI_BSC') IS NOT NULL
	DROP VIEW dbo.VistaRelazioniCFVdaDocZSI_BSC
GO

CREATE VIEW [dbo].[VistaRelazioniCFVdaDocZSI_BSC] AS 
SELECT w.CODCLIFOR
	, W.CODART
	, isnull(isnull(V1.ARTICOLOBSC,X.ARTICOLOBSC),w.codart) AS ARTICOLOBSC
	, isnull(isnull(V1.DSCART_ALT,X.DSCART_ALT),'') AS DSCART_ALT
FROM 
	(	select codclifor,codart
		from  TESTEDOCUMENTI t JOIN RIGHEDOCUMENTI r ON t.PROGRESSIVO = r.IDTESTA
		WHERE t.TIPODOC IN (SELECT p.CODICE FROM PARAMETRIDOC p
							WHERE p.TIPO = 'b' AND p.CLIFOR = 'c'
							AND LEFT(p.CODICE, 1) = 'D'
							AND p.CAUSALEMAG <> 510
							)
		AND codart <> ''
		AND t.ESERCIZIO = year(getdate()) - 2
		GROUP BY codclifor,codart ) AS w 
	LEFT OUTER JOIN  VISTARELAZIONICFV V1 ON v1.codclifor = W.codclifor AND v1.CodArticolo = w.codart 
	LEFT OUTER JOIN VISTARELAZIONICFV X ON X.CODCLIFOR = 'C' AND X.CodArticolo = w.codart
GO

GRANT ALL ON VistaRelazioniCFVdaDocZSI_BSC TO metodo98
GO



IF OBJECT_ID ('dbo.VISTA_EMAILAGENTICLIENTE') IS NOT NULL
	DROP VIEW dbo.VISTA_EMAILAGENTICLIENTE
GO

create view VISTA_EMAILAGENTICLIENTE as 
SELECT 
	A.CODCONTO
	, REPLACE(ISNULL(G1.TELEX, '') + ';' + ISNULL(G2.TELEX, '') + ';' + ISNULL(G3.TELEX, ''), ';;', ';') AS EMAIL_AGENTI
FROM 
	ANAGRAFICARISERVATICF A 
	LEFT OUTER JOIN ANAGRAFICAAGENTI G1 ON A.CODAGENTE1 = G1.CODAGENTE
	LEFT OUTER JOIN ANAGRAFICAAGENTI G2 ON A.CODAGENTE2 = G2.CODAGENTE
	LEFT OUTER JOIN ANAGRAFICAAGENTI G3 ON A.CODAGENTE3 = G3.CODAGENTE
WHERE A.ESERCIZIO = (SELECT MAX(TABESERCIZI.CODICE) FROM TABESERCIZI) -- YWHERE A.ESERCIZIO = 2014 -- YEAR(GETDATE())
	AND (G1.TELEX <> '' OR G2.TELEX <> '' OR G3.TELEX <> '')


GO

GRANT SELECT ON dbo.VISTA_EMAILAGENTICLIENTE TO Metodo98
GO



IF OBJECT_ID ('dbo.fn_GetEMAILCONTATTICLIENTE') IS NOT NULL
	DROP FUNCTION dbo.fn_GetEMAILCONTATTICLIENTE
GO

create function [dbo].[fn_GetEMAILCONTATTICLIENTE](@RIFCONTO VARCHAR(7)) returns varchar(8000) as begin
	declare @Str varchar(8000)
	
	set @Str = ''
	
	select @Str = @Str + (case @Str when '' then '' else ';' end) + R.EMAIL
	from TABELLAPERSONALE R 
	WHERE R.RIFCODCONTO = @RIFCONTO AND  R.CODRUOLO IN (1, 2, 3)
	group by R.EMAIL

	IF @str = ''
	BEGIN
		-- imposto a 0 altrimenti la pubblicazione va in errore 
		SET @str = ''
	END

	return @Str
end
GO

GRANT EXECUTE ON dbo.fn_GetEMAILCONTATTICLIENTE TO Metodo98
GO
GRANT REFERENCES ON dbo.fn_GetEMAILCONTATTICLIENTE TO Metodo98
GO


IF OBJECT_ID ('dbo.ITA_TABREGISTRONOTIFICHEBSC') IS NOT NULL
	DROP TABLE dbo.ITA_TABREGISTRONOTIFICHEBSC
GO

CREATE TABLE dbo.ITA_TABREGISTRONOTIFICHEBSC
	(
	CODCLIFOR           VARCHAR (7) NOT NULL,
	CODART              VARCHAR (50) NOT NULL,
	ARTICOLOBSC         VARCHAR (50) NOT NULL,
	CODLINGUA           VARCHAR (10) NOT NULL,
	IDOBJECT_CLI        INT NOT NULL,
	IDOBJECT_ART        INT NOT NULL,
	IDOBJECT_BOL        INT NOT NULL,
	IDOBJECT_SCH        INT NOT NULL,
	RAGIONESOCIALE      VARCHAR (255),
	DESCRIZIONEARTICOLO VARCHAR (255),
	EMAIL_CLIENTE       VARCHAR (1000),
	EMAIL_AGENTE        VARCHAR (1000),
	IDDOC_BOL               INT NOT NULL,
	FILENAME_BOL            VARCHAR (255),
	CODICEDOCUMENTO_BOL     VARCHAR (255),
	DATAULTIMOINVIO_BOL     DATETIME NULL,
	IDDOC_SCH               INT NOT NULL,
	FILENAME_SCH            VARCHAR (255),
	CODICEDOCUMENTO_SCH     VARCHAR (255),
	DATAULTIMOINVIO_SCH     DATETIME NULL,
	STATOINVIO_BOL SMALLINT,
	STATOINVIO_SCH SMALLINT,
	DSCART_ALT VARCHAR (255),
	PRIMARY KEY (CODCLIFOR, CODART)
	WITH (FILLFACTOR = 90)
	)
GO

GRANT DELETE ON dbo.ITA_TABREGISTRONOTIFICHEBSC TO Metodo98
GO
GRANT INSERT ON dbo.ITA_TABREGISTRONOTIFICHEBSC TO Metodo98
GO
GRANT REFERENCES ON dbo.ITA_TABREGISTRONOTIFICHEBSC TO Metodo98
GO
GRANT SELECT ON dbo.ITA_TABREGISTRONOTIFICHEBSC TO Metodo98
GO
GRANT UPDATE ON dbo.ITA_TABREGISTRONOTIFICHEBSC TO Metodo98
GO


IF OBJECT_ID ('dbo.ITA_SP_UPDATE_REGISTRONOTIFICHEBSC') IS NOT NULL
	DROP PROCEDURE dbo.ITA_SP_UPDATE_REGISTRONOTIFICHEBSC
GO

CREATE PROCEDURE [dbo].[ITA_SP_UPDATE_REGISTRONOTIFICHEBSC]

AS

BEGIN

SET NOCOUNT ON

	INSERT INTO dbo.ITA_TABREGISTRONOTIFICHEBSC (CODCLIFOR, CODART, ARTICOLOBSC, CODLINGUA
	, IDOBJECT_CLI, IDOBJECT_ART, IDOBJECT_BOL, IDOBJECT_SCH
	, RAGIONESOCIALE, DESCRIZIONEARTICOLO
	, EMAIL_CLIENTE, EMAIL_AGENTE
	, IDDOC_BOL
	, FILENAME_BOL, CODICEDOCUMENTO_BOL, DATAULTIMOINVIO_BOL
	, IDDOC_SCH
	, FILENAME_SCH, CODICEDOCUMENTO_SCH, DATAULTIMOINVIO_SCH
	, STATOINVIO_BOL, STATOINVIO_SCH, DSCART_ALT)
 	SELECT CODCLIFOR, CODART, ARTICOLOBSC
		, '' AS CODLINGUA
		, 0 AS IDOBJECT_CLI
		, 0 AS IDOBJECT_ART
		, 0 AS IDOBJECT_BOL
		, 0 AS IDOBJECT_SCH
		, '' AS RAGIONESOCIALE
		, '' AS DESCRIZIONEARTICOLO
		, '' AS EMAIL_CLIENTE
		, '' AS EMAIL_AGENTE
		, 0 AS IDDOC_BOL
		, '' AS FILENAME_BOL
		, '' AS CODICEDOCUMENTO_BOL
		, NULL AS DATAULTIMOINVIO_BOL		
		, 0 AS IDDOC_SCH
		, '' AS FILENAME_SCH
		, '' AS CODICEDOCUMENTO_SCH
		, NULL AS DATAULTIMOINVIO_SCH
		, 0 AS STATOINVIO_BOL
		, 0 AS STATOINVIO_SCH
		, DSCART_ALT
	FROM VistaRelazioniCFVdaDocZSI_BSC
	WHERE NOT EXISTS(SELECT 1 FROM ITA_TABREGISTRONOTIFICHEBSC X WHERE X.CODCLIFOR = VistaRelazioniCFVdaDocZSI_BSC.CODCLIFOR
	AND X.CODART = VistaRelazioniCFVdaDocZSI_BSC.CODART)

	UPDATE 
		ITA_TABREGISTRONOTIFICHEBSC
	SET 
		IDOBJECT_CLI = isnull((SELECT TOP 1 i.IdPubblicazioneKnos FROM ITA_ARCHIVIO_CLIFOR i WHERE i.CodConto = CODCLIFOR), 0)

	UPDATE 
		ITA_TABREGISTRONOTIFICHEBSC
	SET 
		IDOBJECT_ART = isnull((SELECT TOP 1 i.IdPubblicazioneKnos FROM ITA_ARCHIVIO_ARTICOLI i WHERE i.CODICE = ARTICOLOBSC), 0)

	UPDATE 
		ITA_TABREGISTRONOTIFICHEBSC
	SET 
		DESCRIZIONEARTICOLO = (SELECT TOP 1 a.DESCRIZIONE FROM ANAGRAFICAARTICOLI A WHERE CODICE = ARTICOLOBSC)
	
	
	UPDATE  T 
	SET 
		T.RAGIONESOCIALE = a.DSCCONTO1,
		T.CODLINGUA = ISNULL(TN.DESCRIZIONELINGUA9, 'IT-')
	FROM 
		ITA_TABREGISTRONOTIFICHEBSC T JOIN ANAGRAFICACF A ON A.CODCONTO = T.CODCLIFOR
		JOIN TABNAZIONI TN ON A.CODNAZIONE = TN.CODICE
		
		
	UPDATE T 
	SET 
		T.IDDOC_BOL = ISNULL(bol.IDDOC, 0)
		, T.FILENAME_BOL = bol.FILENAME
		, T.IDOBJECT_BOL = bol.IDOBJECT_BOL
	FROM
		ITA_TABREGISTRONOTIFICHEBSC T JOIN Knos_ZSI.DBO.Metodo_View_LinkageBollettini bol 
		ON bol.IDOBJECT_ART = T.IDOBJECT_ART and bol.IDOBJECT_CLIFOR = 38750
	--WHERE
	--	t.IDDOC_BOL = 0
		
				
	UPDATE T 
	SET 
		T.IDDOC_BOL = ISNULL(bol.IDDOC, 0)
		, T.FILENAME_BOL = bol.FILENAME
		, T.IDOBJECT_BOL = bol.IDOBJECT_BOL
	FROM
		ITA_TABREGISTRONOTIFICHEBSC T JOIN Knos_ZSI.DBO.Metodo_View_LinkageBollettini bol 
		ON bol.IDOBJECT_ART = T.IDOBJECT_ART and bol.IDOBJECT_CLIFOR = T.IDOBJECT_CLI



		
	UPDATE 
		ITA_TABREGISTRONOTIFICHEBSC
	SET 
		EMAIL_CLIENTE = [dbo].[fn_GetEMAILCONTATTICLIENTE](CODCLIFOR)
		
	
	UPDATE  T 
	SET 
		T.EMAIL_AGENTE = V.EMAIL_AGENTI
	FROM
		ITA_TABREGISTRONOTIFICHEBSC T JOIN VISTA_EMAILAGENTICLIENTE V
		ON T.CODCLIFOR = V.CODCONTO
			
			
RETURN

END


GO

GRANT EXECUTE ON dbo.ITA_SP_UPDATE_REGISTRONOTIFICHEBSC TO Metodo98
GO




EXEC ITA_SP_UPDATE_REGISTRONOTIFICHEBSC







IF OBJECT_ID ('dbo.ZS_VISTA_NOTIFICHEBSC') IS NOT NULL
	DROP VIEW dbo.ZS_VISTA_NOTIFICHEBSC
GO

create view ZS_VISTA_NOTIFICHEBSC as 
SELECT 
	(CODCLIFOR + ' - ' + RAGIONESOCIALE) AS CLIENTE,
	CODART,
	DESCRIZIONEARTICOLO,
	DSCART_ALT AS DESCRIZIONE_ALTERNATIVA,
	IDOBJECT_CLI,
	IDOBJECT_ART,
	IDOBJECT_BOL,
	IDOBJECT_SCH,
	EMAIL_CLIENTE,
	EMAIL_AGENTE,
	IDDOC_BOL,
	FILENAME_BOL,
	CODICEDOCUMENTO_BOL,
	DATAULTIMOINVIO_BOL,
	IDDOC_SCH,
	FILENAME_SCH,
	CODICEDOCUMENTO_SCH,
	DATAULTIMOINVIO_SCH,
	ARTICOLOBSC,
	CODLINGUA,
	CODCLIFOR,
	RAGIONESOCIALE,
	STATOINVIO_BOL,
	STATOINVIO_SCH
FROM dbo.ITA_TABREGISTRONOTIFICHEBSC

GO

GRANT SELECT ON dbo.ZS_VISTA_NOTIFICHEBSC TO Metodo98
GO

SELECT * FROM ZS_VISTA_NOTIFICHEBSC