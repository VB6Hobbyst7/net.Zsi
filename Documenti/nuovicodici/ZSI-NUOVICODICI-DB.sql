IF OBJECT_ID ('dbo.ZS_GENERAARTICOLO_DB') IS NOT NULL
	DROP PROCEDURE dbo.ZS_GENERAARTICOLO_DB
GO

CREATE PROCEDURE ZS_GENERAARTICOLO_DB(@CODART VARCHAR(50)) AS
BEGIN
	DECLARE @p AS DECIMAL(10)
 	DECLARE @a VARCHAR(50)
 	DECLARE @V VARCHAR(50)
 	
	DECLARE rSqlA CURSOR LOCAL KEYSET FOR 
	SELECT D.PROGRESSIVO, D.artcomposto, D.VERSIONEDBA 
	FROM DISTINTAARTCOMPOSTI d 
	WHERE PATINDEX('%K#%', D.ARTCOMPOSTO) = 0 
	AND LEN(ARTCOMPOSTO) = 6
	AND  d.ARTCOMPOSTO = (CASE WHEN @CODART <> '' THEN @CODART ELSE D.ARTCOMPOSTO END)
	
	OPEN rSqlA

	FETCH NEXT from rSqlA INTO @P, @A, @V
	WHILE (@@FETCH_STATUS <> -1)
		BEGIN
			
			PRINT @P
			PRINT @A
			PRINT @V
			
			PRINT 'PRE'
			SELECT * FROM DISTINTABASE d WHERE d.RIFPROGRESSIVO = @p
			
			UPDATE DISTINTABASE SET SOLOSE = '${3}=XXX' 
			--SELECT * FROM DISTINTABASE
			WHERE SOLOSE = '' AND RIFPROGRESSIVO = @p AND NRRIGA = 1
				
			INSERT INTO dbo.DISTINTABASE (RIFPROGRESSIVO, NRRIGA, POSIZIONE, CODARTCOMPONENTE, DESCRIZIONE, SOLOSE, UM, QTA1, OPERATORE, QTA2, CALCOLO, QTACOSTO, VERSIONECOMPONENTE, SVILUPPACOMPONENTE, DISEGNOALLEGATO, NUMCOMPONENTE, SEQASSEMBLAGGIO, NOTECOMPONENTE, LEADTIMEADJ, PROGCICLO, NUMFASECICLO, UTENTEMODIFICA, DATAMODIFICA, UMCOSTO, FORMULAFRONTIERA)
			SELECT @p, NRRIGA, POSIZIONE, CODARTCOMPONENTE, DESCRIZIONE, SOLOSE, UM, QTA1, OPERATORE, QTA2, CALCOLO, QTACOSTO, VERSIONECOMPONENTE, SVILUPPACOMPONENTE, DISEGNOALLEGATO, NUMCOMPONENTE, SEQASSEMBLAGGIO, NOTECOMPONENTE, LEADTIMEADJ, PROGCICLO, NUMFASECICLO, 'trmxxx', getdate(), UMCOSTO, FORMULAFRONTIERA
			FROM DISTINTABASE xx WHERE RIFPROGRESSIVO = 520 and NRRIGA = 2
			AND NOT EXISTS(SELECT 1 FROM DISTINTABASE x WHERE x.RIFPROGRESSIVO = @p AND x.NRRIGA = 2)
		
			INSERT INTO dbo.DISTINTABASE (RIFPROGRESSIVO, NRRIGA, POSIZIONE, CODARTCOMPONENTE, DESCRIZIONE, SOLOSE, UM, QTA1, OPERATORE, QTA2, CALCOLO, QTACOSTO, VERSIONECOMPONENTE, SVILUPPACOMPONENTE, DISEGNOALLEGATO, NUMCOMPONENTE, SEQASSEMBLAGGIO, NOTECOMPONENTE, LEADTIMEADJ, PROGCICLO, NUMFASECICLO, UTENTEMODIFICA, DATAMODIFICA, UMCOSTO, FORMULAFRONTIERA)
			SELECT @p, NRRIGA, POSIZIONE, replace(CODARTCOMPONENTE, '20021#', @a), (SELECT TOP 1 DESCRIZIONE FROM ANAGRAFICAARTICOLI A WHERE A.CODICE = REPLACE(@A, '#', '')), SOLOSE, UM, QTA1, OPERATORE, QTA2, CALCOLO, QTACOSTO, VERSIONECOMPONENTE, SVILUPPACOMPONENTE, DISEGNOALLEGATO, NUMCOMPONENTE, SEQASSEMBLAGGIO, NOTECOMPONENTE, LEADTIMEADJ, PROGCICLO, NUMFASECICLO, 'trmxxx', getdate(), UMCOSTO, FORMULAFRONTIERA
			FROM DISTINTABASE xx WHERE RIFPROGRESSIVO = 520 and NRRIGA = 3
			AND NOT EXISTS(SELECT 1 FROM DISTINTABASE x WHERE x.RIFPROGRESSIVO = @p AND x.NRRIGA = 3)
				
			PRINT 'POST'
			SELECT * FROM DISTINTABASE d WHERE d.RIFPROGRESSIVO = @p
			
			FETCH NEXT from rSqlA INTO  @P, @A, @V
		END

	CLOSE rSqlA
	DEALLOCATE rSqlA
		
	
	
	
END
GO

GRANT EXECUTE ON dbo.ZS_GENERAARTICOLO_DB TO Metodo98
GO



EXEC ZS_GENERAARTICOLO_DB '20395#'

