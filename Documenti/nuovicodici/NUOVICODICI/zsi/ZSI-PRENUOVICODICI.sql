/*DELETE FROM TABVARIANTI


INSERT INTO dbo.TABVARIANTI (TIPOLOGIA, VARIANTE, POSIZIONE, DESCRIZIONE, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z, DESCRIZIONE1, DESCRIZIONE2, DESCRIZIONE3, DESCRIZIONE4, DESCRIZIONE5, DESCRIZIONE6, DESCRIZIONE7, DESCRIZIONE8, DESCRIZIONE9, UTENTEMODIFICA, DATAMODIFICA)
--VALUES ('62', '000', 1, 'Fusto tondo lt. 60', '60', '62000', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'coralli', '23/11/2016 11:36:43')
SELECT LEFT(codice, 2), substring(codice, 3, 3), row_number() OVER(PARTITION BY LEFT(CODICE, 2) ORDER BY CODICE) 
, descrizione
, ''
, codice
, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'trm', getdate()
FROM ANAGRAFICAARTICOLI WHERE CODICE LIKE '62%' AND arttipologia = 0
and charindex('#', codice) = 0

IF NOT EXISTS(SELECT 1 FROM TABVARIANTI WHERE tipologia = '62' AND variante = 'xxx')
BEGIN
INSERT INTO dbo.TABVARIANTI (TIPOLOGIA, VARIANTE, POSIZIONE, DESCRIZIONE, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z, DESCRIZIONE1, DESCRIZIONE2, DESCRIZIONE3, DESCRIZIONE4, DESCRIZIONE5, DESCRIZIONE6, DESCRIZIONE7, DESCRIZIONE8, DESCRIZIONE9, UTENTEMODIFICA, DATAMODIFICA)
VALUES ('62', 'XXX', 52, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'trm1', getdate())
END

*/
/*
select t.CODART, i.codice, i.varianteimballo, * from GESTIONEPREZZI t join GESTIONEPREZZIRIGHE r on t.PROGRESSIVO = r.RIFPROGRESSIVO
join tabimballi i on i.CODICE = r.COD_IMBALLO
where t.INIZIOVALIDITA = t.FINEVALIDITA
order by t.codart
*/



IF OBJECT_ID ('dbo.ITA_CODICISOST') IS NOT NULL
	DROP TABLE dbo.ITA_CODICISOST
GO

CREATE TABLE dbo.ITA_CODICISOST
	(
	SEL            SMALLINT NULL,
	TABELLA        VARCHAR (80) NOT NULL,
	CAMPO          VARCHAR (80) NOT NULL,
	UtenteModifica VARCHAR (25) NOT NULL,
	DataModifica   DATETIME NOT NULL,
	CONSTRAINT PK__ITA_CODI__2773EE88755D243C PRIMARY KEY (TABELLA,CAMPO)
	)
GO

GRANT DELETE ON dbo.ITA_CODICISOST TO Metodo98
GO
GRANT INSERT ON dbo.ITA_CODICISOST TO Metodo98
GO
GRANT REFERENCES ON dbo.ITA_CODICISOST TO Metodo98
GO
GRANT SELECT ON dbo.ITA_CODICISOST TO Metodo98
GO
GRANT UPDATE ON dbo.ITA_CODICISOST TO Metodo98
GO



IF OBJECT_ID ('dbo.ZS_GENERAARTICOLI') IS NOT NULL
	DROP TABLE dbo.ZS_GENERAARTICOLI
GO

CREATE TABLE dbo.ZS_GENERAARTICOLI
	(
	ARTTIPOLOGIA    VARCHAR (50) NULL,
	CODART          VARCHAR (50) NULL,
	CODICE          VARCHAR (10) NOT NULL,
	DESCRIZIONE     VARCHAR (500) NULL,
	VARIANTEIMBALLO VARCHAR (25) NULL,
	NUOVOCODART     VARCHAR (50) NOT NULL,
	UtenteModifica  VARCHAR (25) NOT NULL,
	DataModifica    DATETIME NOT NULL,
	CONSTRAINT PK__ZS_GENER__74CEDBB85D859AAB PRIMARY KEY (NUOVOCODART,CODICE)
	)
GO

GRANT DELETE ON dbo.ZS_GENERAARTICOLI TO Metodo98
GO
GRANT INSERT ON dbo.ZS_GENERAARTICOLI TO Metodo98
GO
GRANT REFERENCES ON dbo.ZS_GENERAARTICOLI TO Metodo98
GO
GRANT SELECT ON dbo.ZS_GENERAARTICOLI TO Metodo98
GO
GRANT UPDATE ON dbo.ZS_GENERAARTICOLI TO Metodo98
GO


IF OBJECT_ID ('dbo.ZS_GENERAARTICOLI_POST') IS NOT NULL
	DROP TABLE dbo.ZS_GENERAARTICOLI_POST
GO

CREATE TABLE dbo.ZS_GENERAARTICOLI_POST
	(
	CODART         VARCHAR (50) NOT NULL,
	NUOVOCODART    VARCHAR (50) NOT NULL,
	UtenteModifica VARCHAR (25) NOT NULL,
	DataModifica   DATETIME NOT NULL,
	CONSTRAINT PK__ZS_GENER__44A43AF5633E7401 PRIMARY KEY (NUOVOCODART,CODART)
	)
GO

GRANT DELETE ON dbo.ZS_GENERAARTICOLI_POST TO Metodo98
GO
GRANT INSERT ON dbo.ZS_GENERAARTICOLI_POST TO Metodo98
GO
GRANT REFERENCES ON dbo.ZS_GENERAARTICOLI_POST TO Metodo98
GO
GRANT SELECT ON dbo.ZS_GENERAARTICOLI_POST TO Metodo98
GO
GRANT UPDATE ON dbo.ZS_GENERAARTICOLI_POST TO Metodo98
GO





IF OBJECT_ID ('dbo.ZS_VISTA_GENERAARTICOLI_MOV') IS NOT NULL
	DROP VIEW dbo.ZS_VISTA_GENERAARTICOLI_MOV
GO

CREATE VIEW ZS_VISTA_GENERAARTICOLI_MOV
AS

SELECT * FROM (
	SELECT DISTINCT 
	(CASE WHEN charindex('#', R.CODART) > 1 THEN LEFT(R.CODART, charindex('#', R.CODART) - 1) ELSE R.CODART END) AS ARTTIPOLOGIA
	, R.CODART, R.CODIMBALLO AS CODICE, I.DESCRIZIONE AS DESCRIMBALLO, I.VARIANTEIMBALLO, V.DESCRIZIONE AS DESCRVARIANTE
	, (CASE WHEN RIGHT(R.CODART, 3) = 'XXX' THEN REPLACE(R.CODART, 'XXX', I.VARIANTEIMBALLO)
	ELSE R.CODART + '#000' + I.VARIANTEIMBALLO
	END) AS NUOVOCODART
FROM RIGHEDOCUMENTI R WITH (NOLOCK) 
JOIN TABIMBALLI I WITH (NOLOCK) ON I.CODICE = R.CODIMBALLO
JOIN TABVARIANTI v WITH (NOLOCK) ON v.VARIANTE = I.varianteimballo
WHERE R.CODART <> '' AND V.TIPOLOGIA = '62'
AND I.VARIANTEIMBALLO <> ''
	) CTE
WHERE ARTTIPOLOGIA BETWEEN '20000' AND '49999'
	AND charindex('#', substring(nuovocodart, charindex('#', nuovocodart) + 1, 10) ) = 0
	--AND NUOVOCODART NOT IN (SELECT CODICE FROM ANAGRAFICAARTICOLI)
GO

GRANT DELETE ON dbo.ZS_VISTA_GENERAARTICOLI_MOV TO Metodo98
GO
GRANT INSERT ON dbo.ZS_VISTA_GENERAARTICOLI_MOV TO Metodo98
GO
GRANT REFERENCES ON dbo.ZS_VISTA_GENERAARTICOLI_MOV TO Metodo98
GO
GRANT SELECT ON dbo.ZS_VISTA_GENERAARTICOLI_MOV TO Metodo98
GO
GRANT UPDATE ON dbo.ZS_VISTA_GENERAARTICOLI_MOV TO Metodo98
GO


IF OBJECT_ID ('dbo.ZS_VISTA_GENERAARTICOLI_MODELLI') IS NOT NULL
	DROP VIEW dbo.ZS_VISTA_GENERAARTICOLI_MODELLI
GO

CREATE view ZS_VISTA_GENERAARTICOLI_MODELLI as 


SELECT R.CODART AS ARTTIPOLOGIA
	, R.CODART, '' AS CODICE
	, 'XXX' AS DESCRIMBALLO
	, 'XXX' AS VARIANTEIMBALLO, 'XXX' AS DESCRVARIANTE
	, (R.CODART + '#000XXX') AS NUOVOCODART 
	, (SELECT TOP 1 A.DESCRIZIONE FROM ANAGRAFICAARTICOLI A WITH (NOLOCK) WHERE A.CODICE = R.CODART) AS DESCRIZIONE
	, 'A' AS TIPO
	FROM (
SELECT X.CODART FROM ZS_VISTA_GENERAARTICOLI_MOV X 
WHERE charindex('#', X.CODART) = 0
GROUP BY X.CODART
HAVING count(*) > 1) R
GO

GRANT SELECT ON dbo.ZS_VISTA_GENERAARTICOLI_MODELLI TO Metodo98
GO


IF OBJECT_ID ('dbo.ZS_VISTA_GENERAARTICOLI') IS NOT NULL
	DROP VIEW dbo.ZS_VISTA_GENERAARTICOLI
GO

CREATE VIEW ZS_VISTA_GENERAARTICOLI
AS
/*
SELECT *, (SELECT TOP 1 A.DESCRIZIONE FROM ANAGRAFICAARTICOLI A WHERE A.CODICE = CTE.CODART OR A.CODICE = CTE.ARTTIPOLOGIA) AS DESCRIZIONE FROM (
	SELECT distinct
	(CASE WHEN charindex('#', g.CODART) > 1 THEN LEFT(g.CODART, charindex('#', g.CODART) - 1) ELSE g.CODART END) AS ARTTIPOLOGIA
	, g.CODART
	, t.CODICE
	, t.varianteimballo
	, (CASE 
		WHEN charindex('#', g.CODART) > 0 AND charindex('?', g.codart) = 0 THEN LEFT(g.CODART, charindex('#', g.CODART) + 3) + T.VARIANTEIMBALLO 
		WHEN charindex('#', g.CODART) > 0 AND charindex('?', g.codart) > 0 THEN LEFT(g.CODART, charindex('#', g.CODART)) + '000' + T.VARIANTEIMBALLO 
		ELSE
			g.CODART + '#000' + T.VARIANTEIMBALLO 
		END) AS NUOVOCODART
		--, g.*
	FROM GESTIONEPREZZI g JOIN GESTIONEPREZZIRIGHE r ON g.PROGRESSIVO = r.RIFPROGRESSIVO
	JOIN tabimballi t ON t.CODICE = r.cod_imballo 
	JOIN TABVARIANTI v ON v.VARIANTE = t.varianteimballo
	WHERE g.INIZIOVALIDITA = g.FINEVALIDITA and t.varianteimballo  <> '' --AND charindex('?', g.codart) = 0
	--AND t.codice NOT IN (100, 101)
	) CTE
WHERE 
	NUOVOCODART NOT IN (SELECT CODICE FROM ANAGRAFICAARTICOLI)
	AND ARTTIPOLOGIA BETWEEN '20000' AND '49999'
	*/


SELECT X.ARTTIPOLOGIA, X.CODART, X.CODICE, X.DESCRIMBALLO, X.VARIANTEIMBALLO, X.DESCRVARIANTE, X.NUOVOCODART
	, (SELECT TOP 1 A.DESCRIZIONE FROM ANAGRAFICAARTICOLI A WITH (NOLOCK) WHERE A.CODICE = X.CODART OR A.CODICE = X.ARTTIPOLOGIA) AS DESCRIZIONE --, '' AS TIPO
	, (SELECT 'B' FROM ZS_VISTA_GENERAARTICOLI_MOV Z WHERE Z.CODART = X.CODART AND LEFT(Z.CODART, 1) <> '2' GROUP BY Z.CODART HAVING COUNT(Z.CODART) = 1 ) AS TIPO

FROM 
	 ZS_VISTA_GENERAARTICOLI_MOV x -- ON  x.ARTTIPOLOGIA = cte.arttipologia AND x.CODIMBALLO = cte.cod_imballo
UNION
SELECT 
	X.ARTTIPOLOGIA, X.CODART, X.CODICE, X.DESCRIMBALLO, X.VARIANTEIMBALLO, X.DESCRVARIANTE, X.NUOVOCODART, X.DESCRIZIONE, X.TIPO 
FROM 
	ZS_VISTA_GENERAARTICOLI_MODELLI X
GO

GRANT DELETE ON dbo.ZS_VISTA_GENERAARTICOLI TO Metodo98
GO
GRANT INSERT ON dbo.ZS_VISTA_GENERAARTICOLI TO Metodo98
GO
GRANT REFERENCES ON dbo.ZS_VISTA_GENERAARTICOLI TO Metodo98
GO
GRANT SELECT ON dbo.ZS_VISTA_GENERAARTICOLI TO Metodo98
GO
GRANT UPDATE ON dbo.ZS_VISTA_GENERAARTICOLI TO Metodo98
GO


create FUNCTION ITA_GetDscVariante 
	(@Articolo VARCHAR(50), 
	 @Tipologia VARCHAR(8))
RETURNS VARCHAR(80)
AS
BEGIN
	DECLARE @VARESPLICITE AS VARCHAR(255)
	DECLARE @STARTSEARCH int
	DECLARE @LENSEARCH int
	DECLARE @CodVar as VARCHAR(8)
	DECLARE @DscVar as VARCHAR(80)
	
	SELECT @CodVar ='', @DscVar = ''
	
	IF @Tipologia <> '' and @Articolo <> ''
		BEGIN
			SELECT TOP 1  @VARESPLICITE = VARESPLICITE FROM ANAGRAFICAARTICOLI WITH (nolock) WHERE CODICE = @ARTICOLO AND NOT(VARESPLICITE IS NULL)
			SELECT @STARTSEARCH = CHARINDEX(@TIPOLOGIA, @VARESPLICITE)+LEN(@TIPOLOGIA)+1
			SELECT @LENSEARCH =CHARINDEX(';',SUBSTRING(@VARESPLICITE, @STARTSEARCH,255))-1
			
			IF @LENSEARCH > 1
			BEGIN
				SELECT @CodVar = SUBSTRING(@VARESPLICITE, @STARTSEARCH, @LENSEARCH)
				SELECT TOP 1 @DscVar = DESCRIZIONE FROM TABVARIANTI WITH (nolock) WHERE TIPOLOGIA = @TIPOLOGIA AND VARIANTE = @CODVAR
			END
		END
		RETURN @DscVar	
END
GO

GRANT EXECUTE ON dbo.ITA_GetDscVariante TO Metodo98
GO
GRANT REFERENCES ON dbo.ITA_GetDscVariante TO Metodo98
GO



IF OBJECT_ID ('dbo.ZS_VISTADESCRARTICOLIDOCUMENTI') IS NOT NULL
	DROP VIEW dbo.ZS_VISTADESCRARTICOLIDOCUMENTI
GO

create view ZS_VISTADESCRARTICOLIDOCUMENTI as 
	
SELECT r.IDTESTA, r.IDRIGA, r.codart, r.DESCRIZIONEART
, dbo.ITA_GetDscVariante(r.codart, '62') AS descrvar62
, replace(r.descrizioneart, dbo.ITA_GetDscVariante(r.codart, '62'), '') AS descrizioneartnovar62
FROM RIGHEDOCUMENTI r WITH (nolock)
GO

GRANT SELECT ON dbo.ZS_VISTADESCRARTICOLIDOCUMENTI TO Metodo98
GO

CREATE  PROCEDURE ITA_FILLCODICISOST AS
	DECLARE @CSQL AS VARCHAR(5000)
	DECLARE @N AS SMALLINT
	DECLARE @NOMECAMPO AS VARCHAR(80)
	DECLARE @NOMETABELLA AS VARCHAR(80)

	DECLARE rSqlA CURSOR LOCAL KEYSET FOR 
	SELECT C.COLUMN_NAME, C.TABLE_NAME
	FROM INFORMATION_SCHEMA.COLUMNS C JOIN INFORMATION_SCHEMA.TABLES T ON T.TABLE_NAME = C.TABLE_NAME
	WHERE TABLE_TYPE <> 'VIEW' AND DATA_TYPE LIKE '%CHAR%' AND CHARACTER_MAXIMUM_LENGTH >= 50 AND 
	LEFT(C.TABLE_NAME,1) <> 'W' AND  LEFT(C.TABLE_NAME,3) <> 'TMP' AND LEFT(C.TABLE_NAME,4) <> 'TEMP' AND
	C.COLUMN_NAME NOT LIKE '%DESCR%' AND C.COLUMN_NAME NOT LIKE '%NOT%' AND C.COLUMN_NAME NOT LIKE '%DSC%' AND
	C.COLUMN_NAME NOT LIKE '%INDIR%' AND C.COLUMN_NAME NOT LIKE '%LOCAL%' AND C.COLUMN_NAME NOT LIKE '%FORMULA%' AND
	C.COLUMN_NAME NOT LIKE '%RAGS%' AND C.COLUMN_NAME NOT LIKE '%CALCOLO%' AND C.COLUMN_NAME <> 'CODICEPRIMARIO' AND
	C.COLUMN_NAME NOT LIKE '%AGE%' AND C.COLUMN_NAME NOT LIKE '%ATT%' AND C.COLUMN_NAME NOT LIKE '%CAMPO%' AND
	C.COLUMN_NAME NOT LIKE '%CAUS%' AND C.COLUMN_NAME NOT LIKE '%DESP%' AND C.COLUMN_NAME NOT LIKE '%DESUM%' AND
	C.COLUMN_NAME NOT LIKE '%NOME%' AND C.COLUMN_NAME NOT LIKE '%TIPI%' AND C.COLUMN_NAME NOT LIKE '%INTESTA%' 
	ORDER BY C.COLUMN_NAME
	OPEN rSqlA

	FETCH NEXT from rSqlA INTO @NOMECAMPO, @NOMETABELLA
	WHILE (@@FETCH_STATUS <> -1)
		BEGIN
		
			SELECT @N = COUNT(*) FROM ITA_CODICISOST WHERE TABELLA = @NOMETABELLA AND CAMPO = @NOMECAMPO
			
			IF @N = 0
			BEGIN

				IF @NOMETABELLA = 'RELAZIONICFV' AND ( @NOMECAMPO = 'ARTICOLO' OR @NOMECAMPO = 'VARIANTI' )

					BEGIN
				
						INSERT INTO ITA_CODICISOST VALUES (1, @NOMETABELLA, @NOMECAMPO, '', getdate() )

					END
				
				ELSE

					BEGIN
						SET @CSQL = 'SELECT DISTINCT [' + @NOMECAMPO + '] FROM ' + @NOMETABELLA + ' WHERE ISNULL([' + @NOMECAMPO + '],'''') <> '''' '
						--PRINT @CSQL
						EXEC (@CSQL)
						IF @@ROWCOUNT > 0
						BEGIN

							SET @CSQL = 'SELECT TOP 1 ''A'' FROM ' + @NOMETABELLA + ' WHERE [' + @NOMECAMPO + '] IN (SELECT CODICE FROM ANAGRAFICAARTICOLI) AND ISNULL([' + @NOMECAMPO + '],'''') <> '''' ' 
							--PRINT @CSQL
							EXEC (@CSQL)

							IF @@ROWCOUNT <> 0
							BEGIN

								INSERT INTO ITA_CODICISOST VALUES (1, @NOMETABELLA, @NOMECAMPO, '', getdate() )

								--PRINT @NOMETABELLA + ' ' + @NOMECAMPO				

							END

						END
					END				
			
			END
	
			FETCH NEXT from rSqlA INTO @NOMECAMPO, @NOMETABELLA
		END

	CLOSE rSqlA
	DEALLOCATE rSqlA
GO

GRANT EXECUTE ON dbo.ITA_FILLCODICISOST TO Metodo98
GO

create PROCEDURE ITA_UPDATECODICISOST  (@OLDCODART VARCHAR(50), @CODIMBALLO VARCHAR(10), @NEWCODART VARCHAR(50) ) AS
	DECLARE @CSQL_S AS VARCHAR(5000)
	DECLARE @CSQL_U AS VARCHAR(5000)
	DECLARE @CSQL_U1 AS VARCHAR(5000)
	DECLARE @CSQL_U2 AS VARCHAR(5000)
	DECLARE @TABELLA AS VARCHAR(500)

	
	SET @CSQL_U1 = 'UPDATE S SET S.CODART' + @NEWCODART +
	' FROM RIGHEDOCUMENTI R JOIN STORICOMAG S ON S.IDTESTA = R.IDTESTA AND S.RIGADOC = R.IDRIGA' +  
	' WHERE R.CODART = @OLDCODART AND R.CODIMBALLO  = ' + @CODIMBALLO

	
	SET @CSQL_U2 = 'UPDATE P SET P.CODARTICOLO = ' + @NEWCODART + 
	' FROM RIGHEDOCUMENTI R JOIN STORICOMAG S ON S.IDTESTA = R.IDTESTA AND S.RIGADOC = R.IDRIGA ' +
	' JOIN STORICOPREZZIARTICOLO P ON P.RIFSTORICOMAG = S.PROGRESSIVO' +
	' WHERE R.CODART = @OLDCODART AND R.CODIMBALLO  = ' + @CODIMBALLO

	IF LEFT(@oldcodart, 1) <> '2' AND (SELECT count(*) FROM ZS_VISTA_GENERAARTICOLI X WHERE X.CODART = @OLDCODART AND TIPO IN ('A', 'B')) = 1
	BEGIN
	
	PRINT 'SERIE 30000, 40000aaaaa'
	
		DECLARE rSqlA CURSOR LOCAL KEYSET FOR 
	
		SELECT TABELLA
		, 'SELECT ' + '*' + ' FROM ' + TABELLA + ' WHERE '+ CAMPO + ' = ''' + @OLDCODART + ''''  AS STRSQL_S
		, 'UPDATE ' + TABELLA + ' SET ' + CAMPO + ' = ''' + @NEWCODART + ''' WHERE '+ CAMPO + ' = ''' + @OLDCODART + '''' AS STRSQL
		--, * 
		--, 'UPDATE ITA_CODICISOST SET SEL = 0 WHERE TABELLA = ''' + TABELLA + ''' AND (SELECT COUNT(*) FROM ' + TABELLA + ' WHERE ' + CAMPO + ' IN (SELECT Z.CODART FROM ZS_GENERAARTICOLI z)) = 0' AS STRSQL
		FROM ITA_CODICISOST i 
		WHERE i.SEL < 2
		
	
	--	ORDER BY C.COLUMN_NAME
		OPEN rSqlA
	
		FETCH NEXT from rSqlA INTO @TABELLA,  @CSQL_S, @CSQL_U
		WHILE (@@FETCH_STATUS <> -1)
			BEGIN
				
				PRINT @TABELLA
				
				IF @TABELLA = 'TABLOTTIRIORDINO'
				BEGIN
					DELETE FROM TABLOTTIRIORDINO WHERE CODART = @NEWCODART AND EXISTS(SELECT TOP 1 1 FROM TABLOTTIRIORDINO X WHERE X.CODART = @OLDCODART)
				END
				
				PRINT @CSQL_U
				EXEC (@CSQL_U)
	
	
				
	
			
				--PRINT @@ROWCOUNT
		
				FETCH NEXT from rSqlA INTO @TABELLA, @CSQL_S, @CSQL_U
			END
	
		CLOSE rSqlA
		DEALLOCATE rSqlA
			
	END


	IF LEFT(@oldcodart, 1) = '2'
	BEGIN
	
		DECLARE rSqlA CURSOR LOCAL KEYSET FOR 
	
		SELECT TABELLA
		, 'SELECT ' + '*' + ' FROM ' + TABELLA + ' WHERE '+ CAMPO + ' = ''' + @OLDCODART + ''''  AS STRSQL_S
		, (CASE 
			WHEN TABELLA = 'GESTIONEPREZZI' THEN
				'UPDATE ' + TABELLA + ' SET CODART = ''' + @OLDCODART + '#??????'', CODARTRIC = ''' + @OLDCODART + '#___%'' WHERE INIZIOVALIDITA = FINEVALIDITA AND '+ CAMPO + ' = ''' + @OLDCODART + ''''
			WHEN TABELLA = 'RIGHEDOCUMENTI' THEN
				'UPDATE ' + TABELLA + ' SET CODART = ''' + @NEWCODART + ''' WHERE '+ CAMPO + ' = ''' + @OLDCODART + ''' AND CODIMBALLO = ''' + @CODIMBALLO + ''''
			ELSE
				'UPDATE ' + TABELLA + ' SET ' + CAMPO + ' = ''' + @NEWCODART + ''' WHERE '+ CAMPO + ' = ''' + @OLDCODART + ''''
			END)  AS STRSQL
		--, * 
		--, 'UPDATE ITA_CODICISOST SET SEL = 0 WHERE TABELLA = ''' + TABELLA + ''' AND (SELECT COUNT(*) FROM ' + TABELLA + ' WHERE ' + CAMPO + ' IN (SELECT Z.CODART FROM ZS_GENERAARTICOLI z)) = 0' AS STRSQL
		FROM ITA_CODICISOST i 
		WHERE I.SEL = 1
	
	--	ORDER BY C.COLUMN_NAME
		OPEN rSqlA
	
		FETCH NEXT from rSqlA INTO @TABELLA,  @CSQL_S, @CSQL_U
		WHILE (@@FETCH_STATUS <> -1)
			BEGIN
				
				PRINT @TABELLA
				
				
				IF @TABELLA = 'RIGHEDOCUMENTI'
				BEGIN
	
					PRINT @CSQL_U2
					--EXEC (@CSQL_U2)
		
					PRINT @CSQL_U1
					--EXEC (@CSQL_U1)
					
				END
				
				
				PRINT @CSQL_U
				--EXEC (@CSQL_U)
	
	
				
	
				
				--PRINT @@ROWCOUNT
		
				FETCH NEXT from rSqlA INTO @TABELLA, @CSQL_S, @CSQL_U
			END
	
		CLOSE rSqlA
		DEALLOCATE rSqlA
		
	END
GO

GRANT EXECUTE ON dbo.ITA_UPDATECODICISOST TO Metodo98
GO

CREATE  PROCEDURE ZS_GENERAARTICOLI_GESTIONEPREZZI(@CODART VARCHAR(50), @OLDCODART VARCHAR(50)) AS
	DECLARE @CSQL AS VARCHAR(5000)
	DECLARE @N AS SMALLINT
	DECLARE @NOMECAMPO AS VARCHAR(80)
	DECLARE @NOMETABELLA AS VARCHAR(80)
	
	DECLARE @PROGRESSIVO AS DECIMAL(10)
	DECLARE @IDRIGA DECIMAL(10)
	DECLARE @PPROG DECIMAL (10)
	DECLARE @PIDRIGA DECIMAL (10)
	DECLARE @IMBALLO AS VARCHAR(10)
	
	DELETE FROM GESTIONEPREZZI WHERE CODART = @CODART AND CODCLIFOR = 'C' AND INIZIOVALIDITA = FINEVALIDITA

	DECLARE rSqlA CURSOR LOCAL KEYSET FOR 
	SELECT G.PROGRESSIVO, R.IDRIGA, r.COD_IMBALLO
	FROM GESTIONEPREZZI g JOIN GESTIONEPREZZIRIGHE r ON g.PROGRESSIVO = r.RIFPROGRESSIVO
	JOIN TABIMBALLI T ON T.CODICE = R.COD_IMBALLO
	WHERE g.INIZIOVALIDITA = g.FINEVALIDITA  --AND charindex('?', g.codart) = 0
	AND g.CODART = @OLDCODART AND T.VARIANTEIMBALLO = RIGHT(@CODART, 3)
	
	
	
	
	OPEN rSqlA

	FETCH NEXT from rSqlA INTO @PROGRESSIVO, @IDRIGA, @IMBALLO
	WHILE (@@FETCH_STATUS <> -1)
		BEGIN
		
			PRINT @PROGRESSIVO
			PRINT @IDRIGA
			PRINT @IMBALLO
			
			SELECT @PPROG = 0

			SELECT TOP 1 @PPROG = X.PROGRESSIVO 
			FROM GESTIONEPREZZI X WHERE X.CODART = @CODART AND X.INIZIOVALIDITA = X.FINEVALIDITA
			
			IF (@PPROG = 0)
			BEGIN

				
				EXECUTE dbo.NUOVOPROGRESSIVO 'GESTIONEPREZZI', 1, @PPROG OUT
				SELECT "@PPROG" = @PPROG
	
				INSERT INTO dbo.GESTIONEPREZZI (PROGRESSIVO, CODGRUPPOPREZZICF, CODCLIFOR, CODART, CODGRUPPOPREZZIMAG, INIZIOVALIDITA, FINEVALIDITA, USANRLISTINO, TIPOARROT, ARROTALIRE, ARROTAEURO, CODARTRIC, UTENTEMODIFICA, DATAMODIFICA, PROGRESSIVOCTR)
				SELECT @PPROG, CODGRUPPOPREZZICF, CODCLIFOR, @CODART, CODGRUPPOPREZZIMAG, INIZIOVALIDITA, FINEVALIDITA, USANRLISTINO, TIPOARROT, ARROTALIRE, ARROTAEURO, @CODART, UTENTEMODIFICA, DATAMODIFICA, PROGRESSIVOCTR
				FROM GESTIONEPREZZI 
				WHERE PROGRESSIVO = @PROGRESSIVO 
				AND NOT EXISTS(SELECT 1 FROM GESTIONEPREZZI X WHERE X.CODART = @CODART AND X.INIZIOVALIDITA = X.FINEVALIDITA) 
	 		
			END
			
	 			   
	 
			EXECUTE dbo.NUOVOPROGRESSIVO 'GESTIONEPREZZIRIGHE', 1, @PIDRIGA OUT
			SELECT "@PIDRIGA" = @PIDRIGA  
					
			INSERT INTO dbo.GESTIONEPREZZIRIGHE (IDRIGA, RIFPROGRESSIVO, NRLISTINO, UM, QTAMINIMA, PREZZO_MAGG, PREZZO_MAGGEURO, SCONTO_UNICO, SCONTO_AGGIUNTIVO, TIPO, UTENTEMODIFICA, DATAMODIFICA, TP_QTASCONTO, TP_QTACOEFF, COD_IMBALLO, QTA_COLLI)
			SELECT @PIDRIGA, @PPROG, NRLISTINO, UM, QTAMINIMA, PREZZO_MAGG, PREZZO_MAGGEURO, SCONTO_UNICO, SCONTO_AGGIUNTIVO, TIPO, UTENTEMODIFICA, DATAMODIFICA, TP_QTASCONTO, TP_QTACOEFF, COD_IMBALLO, QTA_COLLI
			FROM GESTIONEPREZZIRIGHE 
			WHERE RIFPROGRESSIVO = @PROGRESSIVO AND IDRIGA = @IDRIGA 
			--AND NOT EXISTS(SELECT 1 FROM GESTIONEPREZZIRIGHE X WHERE X.CODART = @CODART AND X.COD_IMBALLO = @IMBALLO AND X.INIZIOVALIDITA = X.FINEVALIDITA) 
	  	



			
			
			
			FETCH NEXT from rSqlA INTO @PROGRESSIVO, @IDRIGA, @IMBALLO
		END

	CLOSE rSqlA
	DEALLOCATE rSqlA
GO

GRANT EXECUTE ON dbo.ZS_GENERAARTICOLI_GESTIONEPREZZI TO Metodo98
GO

CREATE  PROCEDURE ZS_GENERAARTICOLO_PRE(@CODART VARCHAR(50)) AS
BEGIN
 
	DECLARE @ARTPADRE VARCHAR(50)
	
	SET @ARTPADRE = LEFT(@CODART, CHARINDEX('#' , @CODART) -1)
	
	-- TRASFORMO PUNTUALE IN TIPOLOGIE A VARIANTI
	UPDATE ANAGRAFICAARTICOLI SET ARTTIPOLOGIA = 1 WHERE CODICE = @ARTPADRE AND ARTTIPOLOGIA = 0
	
	INSERT INTO dbo.TIPOLOGIEARTICOLI (CODICEART, NUMEROTIP, CODTIPOLOGIA, SELVARIANTI, AGGIUNGIDES, UTENTEMODIFICA, DATAMODIFICA)
	SELECT @ARTPADRE, 1, 'PE', 1, 0, 'trm1', getdate() FROM TABDITTE WHERE NOT EXISTS(SELECT 1 FROM TIPOLOGIEARTICOLI x WHERE x.CODICEART = @ARTPADRE AND x.CODTIPOLOGIA = 'PE')

	INSERT INTO dbo.TIPOLOGIEARTICOLI (CODICEART, NUMEROTIP, CODTIPOLOGIA, SELVARIANTI, AGGIUNGIDES, UTENTEMODIFICA, DATAMODIFICA)
	SELECT @ARTPADRE, 2, '62', 1, 0, 'trm1', getdate() FROM TABDITTE WHERE NOT EXISTS(SELECT 1 FROM TIPOLOGIEARTICOLI x WHERE x.CODICEART = @ARTPADRE AND x.CODTIPOLOGIA = '62')


	
	INSERT INTO dbo.VARIANTIARTICOLI (CODICEART, NUMEROTIP, VARIANTE, TIPOLOGIA, SOLOSE, UTENTEMODIFICA, DATAMODIFICA)
	SELECT @ARTPADRE, 1, '000', 'PE', '', 'trm1', GETDATE() FROM TABDITTE
	WHERE NOT EXISTS(SELECT 1 FROM VARIANTIARTICOLI x WHERE x.CODICEART = @ARTPADRE AND VARIANTE = '000' AND tipologia = 'PE')

	
	INSERT INTO dbo.VARIANTIARTICOLI (CODICEART, NUMEROTIP, VARIANTE, TIPOLOGIA, SOLOSE, UTENTEMODIFICA, DATAMODIFICA)
	SELECT DISTINCT  ARTTIPOLOGIA, 2, varianteimballo, '62', '', 'trm', getdate() FROM ZS_VISTA_GENERAARTICOLI
	WHERE ARTTIPOLOGIA = @ARTPADRE AND NOT EXISTS(SELECT 1 FROM VARIANTIARTICOLI x WHERE x.CODICEART = ARTTIPOLOGIA AND VARIANTE = varianteimballo AND tipologia = '62')
	
	UPDATE TIPOLOGIEARTICOLI SET AGGIUNGIDES = 1, SELVARIANTI = 1 WHERE CODTIPOLOGIA = '62' AND CODICEART = @ARTPADRE

END
GO

GRANT EXECUTE ON dbo.ZS_GENERAARTICOLO_PRE TO Metodo98
GO

create PROCEDURE ZS_GENERAARTICOLO_POST(@CODART VARCHAR(50), @OLDCODART VARCHAR(50)) AS
BEGIN
 
	DECLARE @ARTPADRE VARCHAR(50)
	
	SET @ARTPADRE = LEFT(@CODART, CHARINDEX('#' , @CODART) -1)

	IF (NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST WHERE CODART = @OLDCODART AND NUOVOCODART = @CODART))
	BEGIN
		INSERT INTO dbo.ZS_GENERAARTICOLI_POST (CODART, NUOVOCODART, UtenteModifica, DataModifica)
		VALUES (@OLDCODART, @CODART, 'TRM', GETDATE())
		--SELECT @OLDCODART, @CODART, 'TRM', GETDATE() FROM TABDITTE WHERE NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST WHERE CODART = @OLDCODART AND NUOVOCODART = @CODART)
	
		--EXEC ZS_GENERAARTICOLI_GESTIONEPREZZI @CODART, @OLDCODART

	END
END
GO

GRANT EXECUTE ON dbo.ZS_GENERAARTICOLO_POST TO Metodo98
GO


IF NOT EXISTS(SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TABIMBALLI' AND COLUMN_NAME = 'VARIANTEIMBALLOF')
	ALTER TABLE TABIMBALLI ADD VARIANTEIMBALLOF VARCHAR(25)
GO
