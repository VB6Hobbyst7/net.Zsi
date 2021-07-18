/*

SELECT * 
FROM ITA_CODICISOST i 
WHERE I.SEL = 1

SELECT * FROM ANAGRAFICAARTICOLI
WHERE CODICE = '43161'


SELECT * FROM STAMPALISTINIPARTICOLARI

SELECT a.CODICE, a.DESCRIZIONE, z.codice 
FROM ZS_GENERAARTICOLI z JOIN ANAGRAFICAARTICOLI a ON a.CODICE = z.NUOVOCODART
WHERE z.NUOVOCODART LIKE '32261%'


SELECT * FROM ZS_GENERAARTICOLI

*/



IF OBJECT_ID ('dbo.ITA_UPDATECODICISOST') IS NOT NULL
	DROP PROCEDURE dbo.ITA_UPDATECODICISOST
GO

CREATE PROCEDURE ITA_UPDATECODICISOST  (@OLDCODART VARCHAR(50), @CODIMBALLO VARCHAR(10), @NEWCODART VARCHAR(50) ) AS
	DECLARE @CSQL_S AS VARCHAR(5000)
	DECLARE @CSQL_U AS VARCHAR(5000)
	DECLARE @CSQL_U1 AS VARCHAR(5000)
	DECLARE @CSQL_U2 AS VARCHAR(5000)
	DECLARE @CSQL_U3 AS VARCHAR(5000)
	DECLARE @TABELLA AS VARCHAR(500)

	
	SET @CSQL_U1 = 'UPDATE S SET S.CODART' + @NEWCODART +
	' FROM RIGHEDOCUMENTI R JOIN STORICOMAG S ON S.IDTESTA = R.IDTESTA AND S.RIGADOC = R.IDRIGA' +  
	' WHERE R.CODART = @OLDCODART AND R.CODIMBALLO  = ' + @CODIMBALLO
	
	SET @CSQL_U2 = 'UPDATE P SET P.CODARTICOLO = ' + @NEWCODART + 
	' FROM RIGHEDOCUMENTI R JOIN STORICOMAG S ON S.IDTESTA = R.IDTESTA AND S.RIGADOC = R.IDRIGA ' +
	' JOIN STORICOPREZZIARTICOLO P ON P.RIFSTORICOMAG = S.PROGRESSIVO' +
	' WHERE R.CODART = @OLDCODART AND R.CODIMBALLO  = ' + @CODIMBALLO

	SET @CSQL_U3 = 'UPDATE S SET S.CODART' + @NEWCODART +
	' FROM RIGHEDOCUMENTI R JOIN STORICOCONSEGNE S ON S.IDTESTA = R.IDTESTA AND S.RIGADOC = R.IDRIGA' +  
	' WHERE R.CODART = @OLDCODART AND R.CODIMBALLO  = ' + @CODIMBALLO

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

				PRINT @CSQL_U3
				--EXEC (@CSQL_U3)

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
GO

GRANT EXECUTE ON dbo.ITA_UPDATECODICISOST TO Metodo98
GO



	DECLARE @oldcodart VARCHAR(50) 
	DECLARE @CODIMBALLO VARCHAR(10) 
	DECLARE @NEWCODART VARCHAR(50) 
	
	DECLARE rSqlA CURSOR LOCAL KEYSET FOR 

	SELECT CODART, codice, nuovocodart 
	FROM ZS_GENERAARTICOLI 
	WHERE ARTTIPOLOGIA LIKE '32261%' AND CODICE = 206

--	ORDER BY C.COLUMN_NAME
	OPEN rSqlA

	FETCH NEXT from rSqlA INTO @oldcodart, @codimballo, @newcodart
	WHILE (@@FETCH_STATUS <> -1)
		BEGIN
		
			PRINT @oldcodart
			PRINT @codimballo
			PRINT @newcodart
			
			EXEC ITA_UPDATECODICISOST @OLDCODART, @CODIMBALLO, @NEWCODART

			FETCH NEXT from rSqlA INTO @oldcodart, @codimballo, @newcodart

		END

	CLOSE rSqlA
	DEALLOCATE rSqlA







