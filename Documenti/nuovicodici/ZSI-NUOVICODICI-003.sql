IF OBJECT_ID ('dbo.ITA_CODICISOST') IS NOT NULL
	DROP TABLE dbo.ITA_CODICISOST
GO

CREATE TABLE dbo.ITA_CODICISOST
	(
	SEL            SMALLINT,
	TABELLA        VARCHAR (80) NOT NULL,
	CAMPO          VARCHAR (80) NOT NULL,
	UtenteModifica VARCHAR (25) NOT NULL,
	DataModifica   DATETIME NOT NULL,
	PRIMARY KEY (TABELLA, CAMPO)
	WITH (FILLFACTOR = 90)
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


INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, '_EXCEL_RELAZIONIARTICOLIBSC', 'CODICE', '', '2017-02-02 07:22:13.49')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, '_ITA_TAB_DEFPROV', 'MACROARTICOLO', '', '2017-02-02 07:23:09.67')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, '_XXX_', 'CODART', '', '2017-02-02 07:22:08.387')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, '_ZZZZ', 'MACROARTICOLO', '', '2017-02-02 07:23:09.66')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'ANAGRAFICAARTICOLI', 'CODICE', '', '2017-02-02 07:22:13.58')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'ANAGRAFICAARTICOLICOMM', 'BARCODE', '', '2017-02-02 07:22:06.423')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'ANAGRAFICAARTICOLICOMM', 'CODICEART', '', '2017-02-02 07:22:13.733')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'ANAGRAFICAARTICOLIPROD', 'CODICEART', '', '2017-02-02 07:22:13.76')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'AnagraficaCespiti', 'Ubicazione', '', '2017-02-02 07:24:03.09')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'ANAGRAFICALOTTI', 'CODARTICOLO', '', '2017-02-02 07:22:13.08')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'ARTICOLIFATTORICONVERSIONE', 'CODART', '', '2017-02-02 07:22:08.12')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'ARTICOLIUMPREFERITE', 'CODART', '', '2017-02-02 07:22:08.253')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'ARTICOLIUNITAMISURA', 'CODART', '', '2017-02-02 07:22:08.203')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (2, 'BRIO_PREZZIARTICOLIMP', 'CODART', '', '2017-02-02 07:22:07.467')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'CONSEGNEAGENTE', 'CODART', '', '2017-02-02 07:22:08.13')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'CONSEGNEGIAC', 'CODART', '', '2017-02-02 07:22:08.057')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'CONTROPARTARTICOLI', 'CODART', '', '2017-02-02 07:22:08.23')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'DESCRARTICOLI', 'CODICEART', '', '2017-02-02 07:22:13.77')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'DISTINTAARTCOMPOSTI', 'ARTCOMPOSTO', '', '2017-02-02 07:17:57.16')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'DISTINTAARTCOMPOSTI', 'RIFARTICOLO', '', '2017-02-02 07:23:17.373')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'DISTINTABASE', 'CODARTCOMPONENTE', '', '2017-02-02 07:22:12.86')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'ECCEZIONIPROGPROD', 'CODART', '', '2017-02-02 07:22:07.523')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'EXCEL_RELAZIONIARTICOLIBSC', 'CODICE', '', '2017-02-02 07:22:13.353')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'EXTRAMAG', 'CODART', '', '2017-02-02 07:22:07.007')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'EXTRAMAG', 'MACRO_ART', '', '2017-02-02 07:23:09.64')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'FEATUREARTICOLO', 'CodArticolo', '', '2017-02-02 07:22:13.057')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'GESTIONEPREZZI', 'CODART', '', '2017-02-02 07:22:08.477')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'GESTIONEPREZZI', 'CODARTRIC', '', '2017-02-02 07:22:13.223')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'IMPEGNIORDPROD', 'CODART', '', '2017-02-02 07:22:09.027')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'IMPOSTAZIONISTAMPA', 'AVALORE', '', '2017-02-02 07:22:05.98')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'IMPOSTAZIONISTAMPA', 'DAVALORE', '', '2017-02-02 07:22:14.73')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'ITA_ARCHIVIO_ARTICOLI', 'CODICE', '', '2017-02-02 07:22:13.343')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'ITA_TAB_DEFPROV', 'MACROARTICOLO', '', '2017-02-02 07:23:09.65')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'ITA_TAB_DEFPROV_LOG', 'MACROARTICOLO', '', '2017-02-02 07:23:09.68')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'ITA_TABREGISTRONOTIFICHEBSC', 'ARTICOLOBSC', '', '2017-02-02 07:17:57.443')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'ITA_TABREGISTRONOTIFICHEBSC', 'CODART', '', '2017-02-02 07:22:08.293')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'ITA_WPUBBLICA_ARTICOLI', 'Codice', '', '2017-02-02 07:22:13.38')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'LIFOARTICOLI', 'CODICEART', '', '2017-02-02 07:22:13.87')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'LISTAARTICOLICTRLINV', 'ARTICOLO', '', '2017-02-02 07:17:57.263')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'LISTAMOVIMENTICTRLINV', 'ARTICOLO', '', '2017-02-02 07:17:57.273')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'LISTINIARTICOLI', 'CODART', '', '2017-02-02 07:22:09.097')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'PROGPRODUZIONE', 'CODART', '', '2017-02-02 07:22:08.263')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'PROGPRODUZIONE', 'RAGGRUPPA', '', '2017-02-02 07:23:13.43')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'PROGPRODUZIONE', 'RIFERIMENTI', '', '2017-02-02 07:23:19.743')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'RELAZIONICFV', 'ARTICOLO', '', '2017-02-02 07:17:57.373')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'RELAZIONICFV', 'ARTICOLOBSC', '', '2017-02-02 07:17:57.457')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'RELAZIONICFV', 'RIFERIMENTO', '', '2017-02-02 07:23:19.833')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'RELAZIONICFV', 'VARIANTI', '', '2017-02-02 07:24:04.533')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'RIGHEDOCUMENTI', 'CODART', '', '2017-02-02 07:22:08.03')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'RIGHEDOCUMENTI', 'RifRelazioneCF', '', '2017-02-02 07:23:24.707')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'STAMPALISTINIPARTICOLARI', 'CODART', '', '2017-02-02 07:22:09.37')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'STORICOCONSEGNE', 'CODARTICOLO', '', '2017-02-02 07:22:12.907')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'STORICOMAG', 'CODART', '', '2017-02-02 07:22:09.35')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'STORICOPREZZIARTICOLO', 'CODARTICOLO', '', '2017-02-02 07:22:13.02')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'TAB_ARTICOLIADR', 'CODART', '', '2017-02-02 07:22:08.49')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (1, 'TABLOTTIRIORDINO', 'CODART', '', '2017-02-02 07:22:07.48')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'TIPOLOGIEARTICOLI', 'CODICEART', '', '2017-02-02 07:22:13.787')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'TP_ExtraMag', 'CodArt', '', '2017-02-02 07:22:09.133')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'TP_GOI', 'CodArt', '', '2017-02-02 07:22:09.357')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'TP_PREZZIVENDITA', 'CodArticolo', '', '2017-02-02 07:22:13.03')
GO

INSERT INTO dbo.ITA_CODICISOST (SEL, TABELLA, CAMPO, UtenteModifica, DataModifica)
VALUES (0, 'TP_SINONIMI', 'Cod_Sino', '', '2017-02-02 07:22:06.947')
GO



IF OBJECT_ID ('dbo.ITA_UPDATECODICISOST') IS NOT NULL
	DROP PROCEDURE dbo.ITA_UPDATECODICISOST
GO

CREATE PROCEDURE ITA_UPDATECODICISOST  (@OLDCODART VARCHAR(50), @CODIMBALLO VARCHAR(10), @NEWCODART VARCHAR(50) ) AS
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

	IF LEFT(@oldcodart, 1) <> '2' AND (SELECT count(*) FROM ZS_VISTA_GENERAARTICOLI X WHERE X.NUOVOCODART = @NEWCODART AND TIPO IN ('A', 'B')) = 1
	BEGIN
	
		--PRINT 'SERIE 30000, 40000aaaaa'
	
		DECLARE rSqlA CURSOR LOCAL KEYSET FOR 
	
		SELECT TABELLA
		, 'SELECT ' + '*' + ' FROM ' + TABELLA + ' WHERE '+ CAMPO + ' = ''' + @OLDCODART + ''''  AS STRSQL_S
		, 'UPDATE ' + TABELLA + ' SET ' + CAMPO + ' = ''' + @NEWCODART + ''' WHERE '+ CAMPO + ' = ''' + @OLDCODART + '''' AS STRSQL
		--, * 
		--, 'UPDATE ITA_CODICISOST SET SEL = 0 WHERE TABELLA = ''' + TABELLA + ''' AND (SELECT COUNT(*) FROM ' + TABELLA + ' WHERE ' + CAMPO + ' IN (SELECT Z.CODART FROM ZS_GENERAARTICOLI z)) = 0' AS STRSQL
		FROM ITA_CODICISOST i 
		WHERE i.SEL = 1
		
	
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
				
				--PRINT @CSQL_U
				EXEC (@CSQL_U)
	
	
				
	
			
				--PRINT @@ROWCOUNT
		
				FETCH NEXT from rSqlA INTO @TABELLA, @CSQL_S, @CSQL_U
			END
	
		CLOSE rSqlA
		DEALLOCATE rSqlA
			
	END



	IF LEFT(@oldcodart, 1) = '2'
	BEGIN
	
	/*
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
		*/
		RETURN
		
	END
GO

GRANT EXECUTE ON dbo.ITA_UPDATECODICISOST TO Metodo98
GO




IF NOT EXISTS(SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'EXTRAMAG' AND COLUMN_NAME = 'NRPEZZIIMBALLO')
BEGIN
	ALTER TABLE EXTRAMAG ADD NRPEZZIIMBALLO INT
END
GO




CREATE TRIGGER ZS_TU_ANAGRAFICAARTICOLI ON ANAGRAFICAARTICOLI FOR UPDATE AS
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
      
      /*  PARENT "ANAGRAFICADEPOSITI" MUST EXIST WHEN UPDATING A CHILD IN "ANAGRAFICAARTICOLI"  */
      IF UPDATE(NRPEZZIIMBALLO)
      BEGIN
         UPDATE EM
         SET
         	EM.NRPEZZIIMBALLO = I1.NRPEZZIIMBALLO
         FROM EXTRAMAG EM, INSERTED I1
         WHERE EM.CODART = I1.CODICE
      END
      

      RETURN

/*  ERRORS HANDLING  */
ERROR:
    RAISERROR (@ERRMSG, 1, 1)
    ROLLBACK  TRANSACTION
END

GO






IF OBJECT_ID ('dbo.ZS_GENERAARTICOLO_PRE') IS NOT NULL
	DROP PROCEDURE dbo.ZS_GENERAARTICOLO_PRE
GO

CREATE PROCEDURE ZS_GENERAARTICOLO_PRE(@CODART VARCHAR(50)) AS
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


IF OBJECT_ID ('dbo.ZS_GENERAARTICOLI_GESTIONEPREZZI') IS NOT NULL
	DROP PROCEDURE dbo.ZS_GENERAARTICOLI_GESTIONEPREZZI
GO

CREATE PROCEDURE ZS_GENERAARTICOLI_GESTIONEPREZZI(@CODART VARCHAR(50), @OLDCODART VARCHAR(50)) AS
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



IF OBJECT_ID ('dbo.ZS_GENERAARTICOLI_POST') IS NOT NULL
	DROP TABLE dbo.ZS_GENERAARTICOLI_POST
GO

CREATE TABLE dbo.ZS_GENERAARTICOLI_POST
	(
	ARTTIPOLOGIA    VARCHAR (50) NULL,
	CODART          VARCHAR (50) NULL,
	CODICE          VARCHAR (10) NOT NULL,
	DESCRIZIONE     VARCHAR (500) NULL,
	VARIANTEIMBALLO VARCHAR (25) NULL,
	NUOVOCODART     VARCHAR (50) NOT NULL,
	UtenteModifica  VARCHAR (25) NOT NULL,
	DataModifica    DATETIME NOT NULL,
	CONSTRAINT PK__ZS_GENERAARTICOLI_POST PRIMARY KEY (NUOVOCODART,CODICE)
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





IF OBJECT_ID ('dbo.ZS_GENERAARTICOLO_POST') IS NOT NULL
	DROP PROCEDURE dbo.ZS_GENERAARTICOLO_POST
GO

create PROCEDURE ZS_GENERAARTICOLO_POST(@CODART VARCHAR(50), @OLDCODART VARCHAR(50), @CODIMBALLO VARCHAR(50), @REDO SMALLINT = 0) AS
BEGIN
 
	DECLARE @ARTPADRE VARCHAR(50)
	DECLARE @ARTMODELLO VARCHAR(50)
	
	SET @ARTPADRE = LEFT(@CODART, CHARINDEX('#' , @CODART) -1)
	SET @ARTMODELLO = LEFT(@CODART, LEN(@CODART) -3) + 'XXX'
	

	

	
	IF (NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST WHERE CODART = @OLDCODART AND NUOVOCODART = @CODART ) OR @REDO = 1)
	BEGIN
	
		PRINT @ARTPADRE
		PRINT @ARTMODELLO
		PRINT @CODIMBALLO
		
		INSERT INTO dbo.ZS_GENERAARTICOLI_POST (CODART, NUOVOCODART, UtenteModifica, DataModifica, CODICE)
		SELECT @OLDCODART, @CODART, 'TRM', GETDATE(), @CODIMBALLO
		FROM TABDITTE
		WHERE NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST X WHERE X.CODART = @OLDCODART AND X.NUOVOCODART = @CODART)-- AND X.CODICE = CODICE)
		
		--EXEC ZS_GENERAARTICOLI_GESTIONEPREZZI @CODART, @OLDCODART
		
		
		
		DECLARE @NRPEZZIIMBALLO INT
		

		SELECT @NRPEZZIIMBALLO = G1.QTA_COLLI 
		FROM ITA_VIS_PRZPART_IMBLIST  G1
		WHERE G1.PRIORITA > 1 AND
			(G1.CODART = @CODART OR
			G1.CODART = LEFT(@CODART, LEN(@CODART) -3) + '???' OR
			G1.CODART LIKE @ARTPADRE + '%') --'#??????')
		AND G1.COD_IMBALLO = @CODIMBALLO

		-- DATI ANAGRAFICASTANDARD
 		UPDATE a1
		SET 
			a1.DESCRIZIONE = RTRIM(LEFT(a2.descrizione + ' ' + (SELECT TOP 1 TABVARIANTI.DESCRIZIONE FROM TABVARIANTI WHERE TABVARIANTI.VARIANTE = RIGHT(a1.CODICE, 3) AND TIPOLOGIA = '62'), 80))
			, a1.GRUPPO = a2.GRUPPO
			, a1.CATEGORIA = a2.CATEGORIA
			, a1.CODCATEGORIASTAT = a2.CODCATEGORIASTAT
			, a1.PESONETTO = a2.PESONETTO
			, a1.SUPERFICIE = a2.SUPERFICIE
			, a1.CUBATURA = a2.CUBATURA
			, a1.NOMENCLCOMBINATA1 = a2.NOMENCLCOMBINATA1
			, a1.NOMENCLCOMBINATA2 = a2.NOMENCLCOMBINATA2
			, a1.ORIGINEINTRA = a2.ORIGINEINTRA 
			, a1.CODICEARTIMBALLO = a1.CODICEARTIMBALLO
			, a1.NRPEZZIIMBALLO = @NRPEZZIIMBALLO --a2.NRPEZZIIMBALLO
			, a1.RIFERIMIMBALLO = @CODIMBALLO --(SELECT TOP 1 tabimballi.CODICE FROM TABIMBALLI WHERE tabimballi.VARIANTEIMBALLO = RIGHT(a1.CODICE, 3))
			, a1.AGGIORNAMAG = a2.AGGIORNAMAG
			, a1.MOVIMENTAPARTITE = a2.MOVIMENTAPARTITE
			, a1.MOVIMENTAMATRICOLE = a2.MOVIMENTAMATRICOLE
			, a1.CODDEPOSITO = a2.CODDEPOSITO
			, a1.NRTIPRAGGRUPPATE = a2.NRTIPRAGGRUPPATE
			, a1.ARTCONFIGURATO = a2.ARTCONFIGURATO
		--SELECT rtrim(a2.descrizione + substring(a1.descrizione, (SELECT len(a3.DESCRIZIONE) FROM ANAGRAFICAARTICOLI a3 WHERE ARTTIPOLOGIA = 1 AND CODICE = a1.CODICEPRIMARIO)+1, 80)), *
		FROM ANAGRAFICAARTICOLI a1, ANAGRAFICAARTICOLI a2
		WHERE 
			a1.CODICE = @CODART
			AND a2.CODICE = @ARTMODELLO
			AND LEFT(a1.CODICE , 1) = '2'
			
			
		
				
		UPDATE a1
		SET 
			a1.CODIVA = a2.CODIVA
			, SCONTO1 = a2.SCONTO1
			, a1.SCONTO2 = a2.SCONTO2
			, a1.SCONTO3 = a2.SCONTO3
			, a1.GRUPPOPRZPART = a2.GRUPPOPRZPART
			, a1.GRUPPOPRVPART = a2.GRUPPOPRVPART
			, a1.PROVV = a2.PROVV
			, a1.CODICEALT1 = a2.CODICEALT1
			, a1.CODICEALT2 = a2.CODICEALT2
			, a1.SCGENVENDITEITA = a2.SCGENVENDITEITA
			, a1.SCGENVENDITEEST = a2.SCGENVENDITEEST
			, a1.SCGENACQUISTIITA = a2.SCGENACQUISTIITA
			, a1.SCGENACQUISTIEST = a2.SCGENACQUISTIEST
			, a1.INESAURIMENTO = a2.INESAURIMENTO
			, a1.ESAURITO = a2.ESAURITO
			, a1.QTAMINCONS = a2.QTAMINCONS
			, a1.USAPREZZIPART = a2.USAPREZZIPART
			, a1.FlagCauzioni = a2.FlagCauzioni
			, a1.FLGBARCODEGENDAPROCAUTOMSTD = a2.FLGBARCODEGENDAPROCAUTOMSTD
			, a1.EXPORTECOMMERCE = a2.EXPORTECOMMERCE
			, a1.CODIVAINTRA = a2.CODIVAINTRA
		--SELECT *
		FROM ANAGRAFICAARTICOLIcomm a1, ANAGRAFICAARTICOLIcomm a2
		WHERE 
			a1.CODICEART = @CODART
			AND a2.CODICEART = @ARTMODELLO
			AND LEFT(a1.CODICEART , 1) = '2'
			AND a1.ESERCIZIO = a2.esercizio
			AND LEFT(a1.CODICEART , 1) = '2'
		
		
		UPDATE a1
		SET a1.SCORTAMIN = a2.scortamin
			, a1.SCORTAMAX = a2.SCORTAMAX
			, a1.LIVPRODUZIONE = a2.LIVPRODUZIONE
			, a1.RAGGRPRODUZIONE =  a1.RAGGRPRODUZIONE
			, a1.LIVPRODPREC = a2.LIVPRODPREC
			, a1.TIPOGESTIONE = a2.TIPOGESTIONE
			, a1.LIVRIORDINO = a2.LIVRIORDINO
			, a1.PROVENIENZA = A2.PROVENIENZA
			, a1.ARTALTERNATIVO =a1.ARTALTERNATIVO
			, a1.QMINRIORDACQ = a2.QMINRIORDACQ
			, a1.QMAXRIORDACQ = a2.QMAXRIORDACQ
			, a1.QDELTARIORDACQ = a2.QDELTARIORDACQ
			, a1.TAPPRONTACQ = a2.TAPPRONTACQ
			, a1.TAPPROVVACQ = a2.TAPPROVVACQ
			, a1.LOTTORIFACQ = a2.LOTTORIFACQ
			, a1.ARROTLOTTOACQ = a2.ARROTLOTTOACQ
			, a1.FORNPREFACQ = a2.FORNPREFACQ
			, a1.QMINRIORDPROD = a2.QMINRIORDPROD
			, a1.QMAXRIORDPROD = a2.QMAXRIORDPROD
			, a1.QDELTARIORDPROD = a2.QDELTARIORDPROD
			, a1.TAPPRONTPROD = a2.TAPPRONTPROD 
			, a1.TAPPROVVPROD = a2.TAPPROVVPROD
			, a1.LOTTORIFPROD = a2.LOTTORIFPROD
			, a1.ARROTLOTTOPROD = a2.ARROTLOTTOPROD
			, a1.QMINRIORDLAV = a2.QMINRIORDLAV
			, a1.QMAXRIORDLAV = a2.QMAXRIORDLAV
			, a1.QDELTARIORDLAV = a2.QDELTARIORDLAV
			, a1.TAPPRONTLAV = a2.TAPPRONTLAV
			, a1.TAPPROVVLAV = a2.TAPPROVVLAV
			, a1.LOTTORIFLAV = a2.LOTTORIFLAV
			, a1.ARROTLOTTOLAV = a2.ARROTLOTTOLAV 
			, a1.FORNPREFLAV = a2.FORNPREFLAV
			, a1.LOTTORIORDINO = a2.LOTTORIORDINO
			, a1.TIPOPRODUZIONE = a2.TIPOPRODUZIONE
			, a1.FLOORSTOCK = a2.FLOORSTOCK
			, a1.GRUPPOAPPROV = a2.GRUPPOAPPROV
			, a1.COSTOORDINEACQ = a2.COSTOORDINEACQ
			, a1.COSTOORDINELAV = a2.COSTOORDINELAV
			, a1.COSTOORDINEPROD = a2.COSTOORDINEPROD
			, a1.FATTORESCOSTAMENTO = a2.FATTORESCOSTAMENTO
			, a1.TEMPOCOPERTURA = a2.TEMPOCOPERTURA
			, a1.CONSUMOPREVISTO = a2.CONSUMOPREVISTO
			, a1.MADPREVISTO = a2.MADPREVISTO
			, a1.FATTORESICUREZZA = a2.TIPOPRODOTTO
			, a1.TIPOPRODOTTO = a2.TIPOPRODOTTO
			, a1.GESTIONEMATERIALI = a2.GESTIONEMATERIALI
			, a1.FATTORECOMPRESSIONE = a2.FATTORECOMPRESSIONE
			, a1.GRUPPOPREVISIONE = a2.GRUPPOPREVISIONE
			, a1.FORMULAFRONTIERA = a2.FORMULAFRONTIERA
			, a1.LIVELLOSERVIZIO = a2.LIVELLOSERVIZIO
			, a1.FLAGMPS = a2.FLAGMPS
			, a1.FLAGNETTIFICAMPS = a2.FLAGNETTIFICAMPS
			, a1.CODICEMPS = a2.CODICEMPS
			, a1.LOTTOFABBRICAZIONE = a2.LOTTOFABBRICAZIONE
			, a1.UMLOTTOFABBRICAZIONE = a2.UMLOTTOFABBRICAZIONE
			, a1.KS_GGScadenza = a2.KS_GGScadenza
			, a1.INTERVALLOPIANIF = a2.INTERVALLOPIANIF
			, a1.DATAULTANALISI = a2.DATAULTANALISI
			, a1.GGORIZZONTEDISP = a2.GGORIZZONTEDISP
		--SELECT *
		FROM ANAGRAFICAARTICOLIPROD a1, ANAGRAFICAARTICOLIPROD a2
		WHERE 
			a1.CODICEART = @CODART
			AND a2.CODICEART = @ARTMODELLO
			AND LEFT(a1.CODICEART , 1) = '2'
			AND a1.ESERCIZIO = a2.esercizio
			AND LEFT(a1.CODICEART , 1) = '2'

		
		
		
		UPDATE a1
		SET A1.ScontiPremi = A2.ScontiPremi 
			, A1.SpTrasp = a2.SpTrasp
			, a1.Provvi = a2.Provvi
			, a1.Imballi = a2.Imballi
			, a1.Amministrazione = a2.Amministrazione
			, a1.SpVendita = a2.SpVendita
			, a1.SpMagazzino = a2.SpMagazzino
			, a1.SpeseProd = a2.SpeseProd
			, a1.CostoAggiunto = a2.CostoAggiunto
			, a1.Preleva = a2.Preleva
			, a1.TestMateriePrime = a2.TestMateriePrime
			, a1.TOL = a2.TOL
			, a1.FLAG_BOLLETTINO = a2.FLAG_BOLLETTINO
			, a1.SCOSTAMENTO_PRZ = a2.SCOSTAMENTO_PRZ
			, a1.XLS_CHEMDES = a2.XLS_CHEMDES
			, a1.XLS_CASNR = a2.XLS_CASNR
			, a1.XLS_ORGGOODS = a2.XLS_ORGGOODS
			, a1.XLS_ORGFORMULA = a2.XLS_ORGFORMULA
			, a1.XLS_VARORGFORMULA = a2.XLS_VARORGFORMULA 
			, a1.XLS_OPDIVISION = a2.XLS_OPDIVISION
			, a1.XLS_COSTDIRWAGES = a2.XLS_COSTDIRWAGES
			, a1.XLS_COSTPRODOVERHEAD = a2.XLS_COSTPRODOVERHEAD
			, a1.XLS_COSTRAWMAT = a2.XLS_COSTRAWMAT
			, a1.XLS_COSTRAWMAT_VEN = a2.XLS_COSTRAWMAT_VEN
			, a1.FLAG_ADR = a2.FLAG_ADR
			, a1.PATH_SCHEDASICUREZZA_ITA = a2.PATH_SCHEDASICUREZZA_ITA
			, a1.PATH_SCHEDASICUREZZA_EST = a2.PATH_SCHEDASICUREZZA_EST
			, a1.MACRO_ART = a2.MACRO_ART
			, a1.PRV_VAR = a2.PRV_VAR
			, a1.BSCKNOS = a2.BSCKNOS
			, a1.FLAG_RSPO = a2.FLAG_RSPO
			, a1.NRPEZZIIMBALLO = @NRPEZZIIMBALLO
		--SELECT *
		FROM EXTRAMAG a1, EXTRAMAG a2
		WHERE 
			a1.CODART = @CODART
			AND a2.CODART = @ARTMODELLO
			AND LEFT(a1.CODART , 1) = '2'
		
		-- LISTINI
		DELETE A1 
		FROM LISTINIARTICOLI A1 
		WHERE a1.CODART = @CODART 
			AND LEFT(@CODART , 1) = '2'
		
		INSERT INTO dbo.LISTINIARTICOLI (CODART, NRLISTINO, UM, PREZZO, PREZZOEURO, UTENTEMODIFICA, DATAMODIFICA, DeltaIncremento, TP_CodConto, TP_ConsPP, TP_PrezzoPart, TP_PrezzoPartEuro, TP_Scorporo, TP_Sconti, TP_QTASCONTO, TP_QTACOEFF, TP_QTAMO, TP_Abbuono, TP_DataCambio, TP_ValoreCambio, DATAVALIDITA, TP_FormulaSct, PREZZOCALC, TP_ABBUONOEURO)
		SELECT @CODART, T.NRLISTINO, UM, PREZZO, PREZZOEURO, 'TRM', GETDATE(), DeltaIncremento, TP_CodConto, TP_ConsPP, TP_PrezzoPart, TP_PrezzoPartEuro, TP_Scorporo, TP_Sconti, TP_QTASCONTO, TP_QTACOEFF, TP_QTAMO, TP_Abbuono, TP_DataCambio, TP_ValoreCambio, DATAVALIDITA, TP_FormulaSct, PREZZOCALC, TP_ABBUONOEURO
		FROM LISTINIARTICOLI t 
		WHERE T.CODART = @ARTMODELLO
			AND LEFT(@CODART , 1) = '2'
			AND NOT EXISTS (SELECT 1 FROM LISTINIARTICOLI x WHERE @CODART = X.CODART AND x.NRLISTINO = T.NRLISTINO)
		
		
		
		DELETE A1 
		FROM CONTROPARTARTICOLI A1 
		WHERE a1.CODART = @CODART 
		
		INSERT INTO dbo.CONTROPARTARTICOLI (CODART, ESERCIZIO, NUMERO, SCGEN, UTENTEMODIFICA, DATAMODIFICA)
		SELECT @CODART, T.ESERCIZIO, T.NUMERO, T.SCGEN, 'TRM', GETDATE()
		FROM CONTROPARTARTICOLI t 
		WHERE T.CODART = @ARTMODELLO
			AND LEFT(@CODART , 1) = '2'



		IF LEFT(@CODART , 1) IN ('3', '4')
		BEGIN
			EXEC ITA_UPDATECODICISOST  @OLDCODART, @CODIMBALLO, @CODART    
			
			
			UPDATE a1
			SET 
				a1.NRPEZZIIMBALLO = @NRPEZZIIMBALLO --a2.NRPEZZIIMBALLO
				, a1.RIFERIMIMBALLO = @CODIMBALLO --(SELECT TOP 1 tabimballi.CODICE FROM TABIMBALLI WHERE tabimballi.VARIANTEIMBALLO = RIGHT(a1.CODICE, 3))
			FROM ANAGRAFICAARTICOLI a1
			WHERE 
				a1.CODICE = @CODART
				AND LEFT(a1.CODICE , 1) IN ('3', '4')	 
			
			UPDATE EXTRAMAG
			SET NRPEZZIIMBALLO = @NRPEZZIIMBALLO	
				
		END
						
		-- inserimento in tabella articoli adr
		INSERT INTO dbo.TAB_ARTICOLIADR (CODART, CLASSEADR, NUMORD, NUMONU, NUMPER, GI, UTENTEMODIFICA, DATAMODIFICA, DESIGN_TRASP, COD_GALLERIE, ESENZIONE_QTA)
		SELECT @CODART, CLASSEADR, NUMORD, NUMONU, NUMPER, GI, 'trm', getdate(), DESIGN_TRASP, COD_GALLERIE, ESENZIONE_QTA
		FROM TAB_ARTICOLIADR t 
		WHERE (t.CODART = @ARTMODELLO OR t.CODART = @OLDCODART)
		AND @CODART NOT IN (SELECT x.codart FROM TAB_ARTICOLIADR x)		

	END
END

GO

GRANT EXECUTE ON dbo.ZS_GENERAARTICOLO_POST TO Metodo98
GO
