ALTER function [dbo].[fn_GeTCRMCOLUMNS](@V AS VARCHAR(255)) returns varchar(8000) as begin
	declare @Str varchar(8000)
	
	set @Str = ''
	
	select @Str = @Str + (case @Str when '' then '' else ';' end) + I.COLUMN_NAME --+ ''''  + I.COLUMN_NAME + ''' AS ' + I.COLUMN_NAME
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


IF OBJECT_ID ('dbo.EXCEL_EXPORT2CRM_CLIENTI') IS NOT NULL
	DROP VIEW dbo.EXCEL_EXPORT2CRM_CLIENTI
GO

create view EXCEL_EXPORT2CRM_CLIENTI as 

SELECT 
	(CASE WHEN LEFT(DSCCONTO1, 4) = 'ZZZZ' THEN 0 ELSE 1 END) AS STATOANAGRAFICA,
	--a.TIPOCONTO,
	a.CODCONTO AS CODICEAZIENDA,
	a.DSCCONTO1,
	a.DSCCONTO2,
	--a.CODMASTRO,
	a.INDIRIZZO,
	a.CAP,
	a.LOCALITA,
	a.PROVINCIA,
	a.TELEFONO AS TELEFONO_CLI,
	a.FAX AS FAX_CLI,
	a.TELEX AS EMAIL_cliente,
	a.CODFISCALE,
	a.PARTITAIVA,
	--a.CODICEISO,
	replace(replace(a.NOTE, CHAR(13), ''), CHAR(10), '') AS NOTE_CLI,
	a.INDIRIZZOINTERNET
	--a.CODNAZIONE
	, n1.DESCRIZIONE AS DESCR_NAZIONE
	--, r.CODPAG, 
	, p.DESCRIZIONE AS DESCR_PAGAMENTO
	--, R.PORTO, 
	, tp.DESCRIZIONE AS DESCR_PORTO
	, R.CODAGENTE1
	, AA.DSCAGENTE
	--, R.CODSETTORE
	, TS.DESCRIZIONE AS DESCR_SETTORE
	, E.FUNZIONARIO
	, EC.Gruppo, EC.Cosmetica, EC.Household, EC.Industrial_applications
  FROM ANAGRAFICACF a 
JOIN ANAGRAFICARISERVATICF r with (nolock) ON a.CODCONTO = r.CODCONTO 
JOIN EXTRACLIENTI E with (nolock) ON E.CODCONTO = A.CODCONTO
JOIN EXTRACLIENTICRM EC with (nolock) ON EC.CODCONTO = A.CODCONTO
LEFT OUTER JOIN TABPORTO tp with (nolock) ON tp.CODICE = r.PORTO
LEFT OUTER JOIN tabnazioni n1 with (nolock) ON n1.CODICE = a.CODNAZIONE
LEFT OUTER JOIN TABPAGAMENTI p with (nolock) ON p.CODICE = r.CODPAG
LEFT OUTER JOIN ANAGRAFICAAGENTI AA with (nolock) ON AA.CODAGENTE = r.CODAGENTE1
LEFT OUTER JOIN TABSETTORI TS with (nolock) ON TS.CODICE = R.CODSETTORE

WHERE a.TIPOCONTO = 'c' AND r.ESERCIZIO = year(getdate())
GO

GRANT SELECT ON dbo.EXCEL_EXPORT2CRM_CLIENTI TO Metodo98
GO




