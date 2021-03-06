IF OBJECT_ID ('dbo.ZS_TABVARIANTI') IS NOT NULL
	DROP TABLE dbo.ZS_TABVARIANTI
GO

CREATE TABLE dbo.ZS_TABVARIANTI
	(
	SEL	SMALLINT,
	TIPOLOGIA      VARCHAR (3) NOT NULL,
	VARIANTE       VARCHAR (8) NOT NULL,
	POSIZIONE      SMALLINT,
	DESCRIZIONE    VARCHAR (80),
	UTENTEMODIFICA VARCHAR (25) NOT NULL,
	DATAMODIFICA   DATETIME NOT NULL,
	CONSTRAINT PK_ZS_TABVARIANTI PRIMARY KEY (TIPOLOGIA, VARIANTE)
	WITH (FILLFACTOR = 90)
	)
GO


GRANT DELETE ON dbo.ZS_TABVARIANTI TO Metodo98
GO
GRANT INSERT ON dbo.ZS_TABVARIANTI TO Metodo98
GO
GRANT REFERENCES ON dbo.ZS_TABVARIANTI TO Metodo98
GO
GRANT SELECT ON dbo.ZS_TABVARIANTI TO Metodo98
GO
GRANT UPDATE ON dbo.ZS_TABVARIANTI TO Metodo98
GO


INSERT INTO dbo.ZS_TABVARIANTI (SEL, TIPOLOGIA, VARIANTE, POSIZIONE, DESCRIZIONE, UTENTEMODIFICA, DATAMODIFICA)
SELECT 0, TIPOLOGIA, VARIANTE, POSIZIONE, DESCRIZIONE, UTENTEMODIFICA, GETDATE()
FROM TABVARIANTI 
WHERE TIPOLOGIA = '62'
GO


/*
IF OBJECT_ID ('dbo.ZS_VISTA_GENERAARTICOLI') IS NOT NULL
	DROP VIEW dbo.ZS_VISTA_GENERAARTICOLI
GO

CREATE VIEW ZS_VISTA_GENERAARTICOLI
AS
SELECT * FROM (
SELECT X.ARTTIPOLOGIA, X.CODART, X.CODICE, X.DESCRIMBALLO, X.VARIANTEIMBALLO, X.DESCRVARIANTE, X.NUOVOCODART
	, (SELECT TOP 1 A.DESCRIZIONE FROM ANAGRAFICAARTICOLI A WITH (NOLOCK) WHERE A.CODICE = X.CODART OR A.CODICE = X.ARTTIPOLOGIA) AS DESCRIZIONE --, '' AS TIPO
	, (SELECT 'B' FROM ZS_VISTA_GENERAARTICOLI_MOV Z WHERE Z.CODART = X.CODART AND LEFT(Z.CODART, 1) <> '2' GROUP BY Z.CODART HAVING COUNT(Z.CODART) = 1 ) AS TIPO
	, (CASE WHEN X.CODICE <> 'XXX' THEN 0 ELSE 9999 END) AS ORDINAMENTO

FROM 
	 ZS_VISTA_GENERAARTICOLI_MOV x -- ON  x.ARTTIPOLOGIA = cte.arttipologia AND x.CODIMBALLO = cte.cod_imballo
WHERE
	NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST XX WHERE XX.CODART = X.CODART AND XX.NUOVOCODART = X.NUOVOCODART)
UNION
SELECT 
	X.ARTTIPOLOGIA, X.CODART, X.CODICE, X.DESCRIMBALLO, X.VARIANTEIMBALLO, X.DESCRVARIANTE, X.NUOVOCODART, X.DESCRIZIONE, X.TIPO 
	, (CASE WHEN X.CODICE <> 'XXX' THEN 0 ELSE 9999 END) AS ORDINAMENTO
FROM 
	ZS_VISTA_GENERAARTICOLI_MODELLI X
WHERE 
	NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST XX WHERE XX.CODART = X.CODART AND XX.NUOVOCODART = X.NUOVOCODART)

UNION

SELECT LEFT(A.CODICE, 5) AS ARTTIPOLOGIA
	, a.CODICE AS CODART
	, 'XXX' AS CODICE
	, '' AS DESCRIMBALLO
	, 'XXX' AS VARIANTEIMBALLO
	, '' AS DESCRVARIANTE
	, A.CODICE + '#000000XXX' AS NUOVOCODART
	, ''
	, 'X' AS TIPO
	, 999	
	--, * 
FROM ANAGRAFICAARTICOLI A
WHERE 
	ARTTIPOLOGIA = 1 AND CODICE IN (SELECT CODICE FROM ZS_IMPORT_INV00)
	AND NOT EXISTS(SELECT 1 FROM ANAGRAFICAARTICOLI X WHERE A.CODICE =  X.CODICE + '#000000XXX')
	AND LEFT(A.CODICE, 2) >= '32'

UNION 





SELECT DISTINCT  LEFT(A.CODICE, 5) AS ARTTIPOLOGIA
	, a.CODICE AS CODART
	, 'XXX' AS CODICE
	, '' AS DESCRIMBALLO
	, 'XXX' AS VARIANTEIMBALLO
	, '' AS DESCRVARIANTE
	, A.CODICE + '#000000XXX' AS NUOVOCODART
	, ''
	, 'X' AS TIPO
	, 999	
	--, * 
FROM ANAGRAFICAARTICOLI A
WHERE 
	ARTTIPOLOGIA = 1 --AND CODICE NOT IN (SELECT CODICE FROM ZS_IMPORT_INV00)
	AND LEFT(A.CODICE, 2) >= '32' and A.CODICE <> '62'
	AND (EXISTS(SELECT TOP 1 1 FROM STORICOMAG XX  WHERE XX.CODART = A.CODICE)
	OR NOT EXISTS(SELECT 1 FROM ANAGRAFICAARTICOLI X WHERE x.CODICE =  a.CODICE + '#000000XXX'))

	
	
UNION

SELECT LEFT(A.CODICE, 5) AS ARTTIPOLOGIA
	, a.CODICE AS CODART
	, I.CODICE
	, I.DESCRIZIONE AS DESCRIMBALLO
	, I.VARIANTEIMBALLO
	, V.DESCRIZIONE AS DESCRVARIANTE
	, REPLACE(A.CODICE, 'XXX', V.VARIANTE) AS NUOVOCODART
	, A.DESCRIZIONE + ' ' + V.DESCRIZIONE
	, 'X' AS TIPO
	, 999	
	--, * 
FROM ANAGRAFICAARTICOLI a, TABIMBALLI I, TABVARIANTI V 
WHERE 
	ARTTIPOLOGIA = 0 AND LEN(A.CODICE) > 12 AND patindex('%K#%', a.codice) = 0 
	AND (LEFT(a.CODICE, 2) = '21' OR LEFT(a.CODICE, 5) = '25102' OR LEFT(a.CODICE, 5) = '25104' OR LEFT(a.CODICE, 5) = '20067')
	AND RIGHT(a.CODICE, 3) = 'XXX' 
	AND V.VARIANTE = I.VARIANTEIMBALLO AND V.TIPOLOGIA = '62'
	AND I.VARIANTEIMBALLO = '396'
	

UNION

SELECT LEFT(A.CODICE, 5) AS ARTTIPOLOGIA
	, a.CODICE AS CODART
	, I.CODICE
	, I.DESCRIZIONE AS DESCRIMBALLO
	, I.VARIANTEIMBALLO
	, V.DESCRIZIONE AS DESCRVARIANTE
	, REPLACE(A.CODICE, 'XXX', V.VARIANTE) AS NUOVOCODART
	, A.DESCRIZIONE + ' ' + V.DESCRIZIONE
	, 'X' AS TIPO
	, 999	
	--, * 
FROM ANAGRAFICAARTICOLI a, TABIMBALLI I, TABVARIANTI V 
WHERE 
	ARTTIPOLOGIA = 0 AND LEN(A.CODICE) > 12 AND patindex('%K#%', a.codice) = 0 
	AND (LEFT(a.CODICE, 2) <> '21' OR LEFT(a.CODICE, 1) IN ('2', '3', '4'))
	AND RIGHT(a.CODICE, 3) = 'XXX' 
	AND V.VARIANTE = I.VARIANTEIMBALLO AND V.TIPOLOGIA = '62'
	AND I.VARIANTEIMBALLO IN ('394', '392', '390')
	AND A.CODICE IN (SELECT DISTINCT RD.CODART FROM RIGHEDOCUMENTI RD WHERE RD.TIPODOC IN ('OCC', 'OCE', 'OCG') AND ESERCIZIO >= 2016)
	
) CTE
WHERE NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST XX WHERE XX.NUOVOCODART = CTE.NUOVOCODART)
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


*/

IF OBJECT_ID ('dbo.ZS_GENERAARTICOLI_EXTRA') IS NOT NULL
	DROP TABLE dbo.ZS_GENERAARTICOLI_EXTRA
GO

CREATE TABLE dbo.ZS_GENERAARTICOLI_EXTRA
	(
	ARTTIPOLOGIA    VARCHAR (50),
	CODART          VARCHAR (50) NOT NULL,
	CODICE          VARCHAR (10) NOT NULL,
	DESCRIZIONE     VARCHAR (500),
	VARIANTEIMBALLO VARCHAR (25),
	DESCRVARIANTE     VARCHAR (500),
	NUOVOCODART     VARCHAR (50) NOT NULL,
	UtenteModifica  VARCHAR (25) NOT NULL,
	DataModifica    DATETIME NOT NULL,
	TIPO VARCHAR(5),
	ORDINAMENTO     INT,
	CONSTRAINT PK__ZS_GENERAARTICOLI_EXTRA PRIMARY KEY (NUOVOCODART, CODICE, CODART)
	)
GO

GRANT DELETE ON dbo.ZS_GENERAARTICOLI_EXTRA TO Metodo98
GO
GRANT INSERT ON dbo.ZS_GENERAARTICOLI_EXTRA TO Metodo98
GO
GRANT REFERENCES ON dbo.ZS_GENERAARTICOLI_EXTRA TO Metodo98
GO
GRANT SELECT ON dbo.ZS_GENERAARTICOLI_EXTRA TO Metodo98
GO
GRANT UPDATE ON dbo.ZS_GENERAARTICOLI_EXTRA TO Metodo98
GO



/*


per tutti gli articoli tranne quelli sotto indicati, creare codice con imballo 394, 392, 390
per gli articoli  21000 creare codice con imballo 396
per gli articoli che vengono venduti in sacchi, creare i codici con imballo 370
per i 32000 NON venduti in sacchi creare codice con imballo 986, 987, 988

*/


INSERT INTO dbo.ZS_GENERAARTICOLI_EXTRA (ARTTIPOLOGIA, CODART, CODICE, DESCRIZIONE, VARIANTEIMBALLO, NUOVOCODART, UtenteModifica, DataModifica, ORDINAMENTO, TIPO, DESCRVARIANTE)
SELECT ARTTIPOLOGIA, CODART, CODICE, DESCRIZIONE, VARIANTEIMBALLO, NUOVOCODART, 'TRMXXX', GETDATE(), ORDINAMENTO, TIPO, DESCRVARIANTE FROM (
-- ****************** per gli articoli che vengono venduti in sacchi, creare i codici con imballo 370
SELECT DISTINCT 
	LEFT(r.CODART, 5) AS ARTTIPOLOGIA
	, r.CODART AS CODART
	, 'XXX' AS CODICE
	, '' AS DESCRIMBALLO
	, I.VARIANTEIMBALLO AS VARIANTEIMBALLO
	, (SELECT TOP 1 x.descrizione FROM TABVARIANTI x WHERE x.TIPOLOGIA = '62' AND x.VARIANTE = i.VARIANTEIMBALLO) AS DESCRVARIANTE
	, LEFT(r.CODART, 5) + '#000000' + I.VARIANTEIMBALLO AS NUOVOCODART
	, '' AS DESCRIZIONE
	, '3' AS TIPO
	, 999 AS ORDINAMENTO	
FROM RIGHEDOCUMENTI r , TABIMBALLI I
WHERE r.CODART <>'' 
AND I.VARIANTEIMBALLO = '370' AND RIGHT(r.CODART, 3) = 'XXX'
AND r.CODIMBALLO IN ('9', '300', '301', '302')
AND esercizio >= 2016 
AND LEFT(r.CODART, 1) IN ('2', '3', '4')
AND NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST z WHERE z.NUOVOCODART = LEFT(r.CODART, 5) + '#000000' + I.VARIANTEIMBALLO)

UNION
/*
-- ****************** per gli articoli che vengono venduti in sacchi, creare i codici con imballo 370
SELECT DISTINCT 
	LEFT(r.CODART, 5) AS ARTTIPOLOGIA
	, r.CODART AS CODART
	, 'XXX' AS CODICE
	, '' AS DESCRIMBALLO
	, 'XXX' AS VARIANTEIMBALLO
	, '' AS DESCRVARIANTE
	, LEFT(r.CODART, 5) + '#000000XXX' AS NUOVOCODART
	, ''  AS DESCRIZIONE
	, '3' AS TIPO
	, 1 AS ORDINAMENTO
FROM RIGHEDOCUMENTI r , TABIMBALLI I
WHERE r.CODART <>'' 
AND I.VARIANTEIMBALLO = '370'
AND r.CODIMBALLO IN ('9', '300', '301', '302')
AND esercizio >= 2016 
AND LEFT(r.CODART, 1) IN ('2', '3', '4')
AND NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST z WHERE z.NUOVOCODART = LEFT(r.CODART, 5) + '#000000XXX')

UNION
*/
-- ************** per i 32000 NON venduti in sacchi creare codice con imballo 986, 987, 988
SELECT DISTINCT 
	LEFT(r.CODART, 5) AS ARTTIPOLOGIA
	, r.CODART AS CODART
	, 'XXX' AS CODICE
	, '' AS DESCRIMBALLO
	, i.VARIANTEIMBALLOF AS VARIANTEIMBALLO
	, (SELECT TOP 1 x.descrizione FROM TABVARIANTI x WHERE x.TIPOLOGIA = '62' AND x.VARIANTE = i.VARIANTEIMBALLOF) AS DESCRVARIANTE
	, LEFT(r.CODART, 5) + '#000000' + I.VARIANTEIMBALLOF AS NUOVOCODART
	, '' AS DESCRIZIONE
	, '4' AS TIPO
	, 999 AS ORDINAMENTO
FROM RIGHEDOCUMENTI r , TABIMBALLI I
WHERE r.CODART <>'' 
AND I.VARIANTEIMBALLOF IN ('986', '987', '988')
AND r.CODIMBALLO NOT IN ('9', '300', '301', '302')
AND esercizio >= 2016 
AND LEFT(r.CODART, 2) ='32'
AND NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST z WHERE z.NUOVOCODART = LEFT(r.CODART, 5) + '#000000' + I.VARIANTEIMBALLOF)

UNION
-- ************** per i 32000 NON venduti in sacchi creare codice con imballo 986, 987, 988
SELECT DISTINCT 
	LEFT(r.CODART, 5) AS ARTTIPOLOGIA
	, r.CODART AS CODART
	, 'XXX' AS CODICE
	, '' AS DESCRIMBALLO
	, 'XXX' AS VARIANTEIMBALLO
	, '' AS DESCRVARIANTE
	, LEFT(r.CODART, 5) + '#000000XXX' AS NUOVOCODART
	, '' AS DESCRIZIONE
	, '4' AS TIPO
	, 1 AS ORDINAMENTO	
FROM RIGHEDOCUMENTI r , TABIMBALLI I
WHERE r.CODART <>'' 
AND I.VARIANTEIMBALLOF IN ('986', '987', '988')
AND r.CODIMBALLO NOT IN ('9', '300', '301', '302')
AND esercizio >= 2016 
AND LEFT(r.CODART, 2) ='32'
AND NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST z WHERE z.NUOVOCODART = LEFT(r.CODART, 5) + '#000000XXX')

UNION

-- ************ per gli articoli  21000 creare codice con imballo 396

SELECT LEFT(A.CODICE, 5) AS ARTTIPOLOGIA
	, a.CODICE AS CODART
	, I.CODICE
	, I.DESCRIZIONE AS DESCRIMBALLO
	, I.VARIANTEIMBALLO
	, V.DESCRIZIONE AS DESCRVARIANTE
	, REPLACE(A.CODICE, 'XXX', V.VARIANTE) AS NUOVOCODART
	, A.DESCRIZIONE + ' ' + V.DESCRIZIONE
	, '2' AS TIPO
	, 999	
	--, * 
FROM ANAGRAFICAARTICOLI a, TABIMBALLI I, TABVARIANTI V 
WHERE 
	ARTTIPOLOGIA = 0 AND LEN(A.CODICE) > 12 AND patindex('%K#%', a.codice) = 0 
	AND (LEFT(a.CODICE, 2) = '21' OR LEFT(a.CODICE, 5) = '25102' OR LEFT(a.CODICE, 5) = '25104' OR LEFT(a.CODICE, 5) = '20067')
	AND RIGHT(a.CODICE, 3) = 'XXX' 
	AND V.VARIANTE = I.VARIANTEIMBALLO AND V.TIPOLOGIA = '62'
	AND I.VARIANTEIMBALLO = '396'
	

	
) CTE


INSERT INTO dbo.ZS_GENERAARTICOLI_EXTRA (ARTTIPOLOGIA, CODART, CODICE, DESCRIZIONE, VARIANTEIMBALLO, NUOVOCODART, UtenteModifica, DataModifica, ORDINAMENTO, TIPO, DESCRVARIANTE)
SELECT LEFT(A.CODICE, 5) AS ARTTIPOLOGIA
	, a.CODICE AS CODART
	, I.CODICE
	, I.DESCRIZIONE AS DESCRIMBALLO
	, I.VARIANTEIMBALLO
	
	, LEFT(a.CODICE, 5) + '#000000' + I.VARIANTEIMBALLO AS NUOVOCODART
	--, A.DESCRIZIONE + ' ' + V.DESCRIZIONE
	
	, 'TRMXXX' AS UTENTEMODIFICA, GETDATE() AS DATAMODIFICA
	, 999	AS ORDINAMENTO
	, '1' AS TIPO
	, V.DESCRIZIONE AS DESCRVARIANTE
	--, * 
FROM ANAGRAFICAARTICOLI a, TABIMBALLI I, TABVARIANTI V 
WHERE LEFT(A.CODICE, 2) <> '21' AND (LEFT(A.CODICE, 1) IN ('2') OR LEFT(A.CODICE, 2) IN ('32', '42')) AND PATINDEX('%k%', A.CODICE) = 0 AND LEN(A.CODICE) = 5 
AND V.VARIANTE = I.VARIANTEIMBALLO AND V.TIPOLOGIA = '62'
AND I.VARIANTEIMBALLO IN ('394', '392', '390')
AND A.CODICE NOT IN (SELECT ZS_GENERAARTICOLI_EXTRA.CODART FROM  ZS_GENERAARTICOLI_EXTRA)

UNION


SELECT DISTINCT 
	LEFT(r.CODART, 5) AS ARTTIPOLOGIA
	, r.CODART AS CODART
	, 'XXX' AS CODICE
	, '' AS DESCRIMBALLO
	, 'XXX' AS VARIANTEIMBALLO
	
	, LEFT(r.CODART, 5) + '#000000XXX' AS NUOVOCODART
	
	, 'TRMXXX' AS UTENTEMODIFICA, GETDATE() AS DATAMODIFICA
	, 999	AS ORDINAMENTO
	, '1' AS TIPO
	, '' AS DESCRVARIANTE
FROM RIGHEDOCUMENTI r , TABIMBALLI I
WHERE r.CODART = '43082' 
AND I.VARIANTEIMBALLO IN ( 'XXX')

AND NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST z WHERE z.NUOVOCODART = LEFT(r.CODART, 5) + '#000000XXX')
	







IF OBJECT_ID ('dbo.ZS_VISTA_GENERAARTICOLI') IS NOT NULL
	DROP VIEW dbo.ZS_VISTA_GENERAARTICOLI
GO

CREATE VIEW ZS_VISTA_GENERAARTICOLI
AS
SELECT * FROM (
SELECT X.ARTTIPOLOGIA, X.CODART, X.CODICE, X.DESCRIMBALLO, X.VARIANTEIMBALLO, X.DESCRVARIANTE, X.NUOVOCODART
	, (SELECT TOP 1 A.DESCRIZIONE FROM ANAGRAFICAARTICOLI A WITH (NOLOCK) WHERE A.CODICE = X.CODART OR A.CODICE = X.ARTTIPOLOGIA) AS DESCRIZIONE --, '' AS TIPO
	, (SELECT 'B' FROM ZS_VISTA_GENERAARTICOLI_MOV Z WHERE Z.CODART = X.CODART AND LEFT(Z.CODART, 1) <> '2' GROUP BY Z.CODART HAVING COUNT(Z.CODART) = 1 ) AS TIPO
	, (CASE WHEN X.CODICE <> 'XXX' THEN 0 ELSE 9999 END) AS ORDINAMENTO

FROM 
	 ZS_VISTA_GENERAARTICOLI_MOV x -- ON  x.ARTTIPOLOGIA = cte.arttipologia AND x.CODIMBALLO = cte.cod_imballo
WHERE
	NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST XX WHERE XX.CODART = X.CODART AND XX.NUOVOCODART = X.NUOVOCODART)
UNION
SELECT 
	X.ARTTIPOLOGIA, X.CODART, X.CODICE, X.DESCRIMBALLO, X.VARIANTEIMBALLO, X.DESCRVARIANTE, X.NUOVOCODART, X.DESCRIZIONE, X.TIPO 
	, (CASE WHEN X.CODICE <> 'XXX' THEN 0 ELSE 9999 END) AS ORDINAMENTO
FROM 
	ZS_VISTA_GENERAARTICOLI_MODELLI X
WHERE 
	NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST XX WHERE XX.CODART = X.CODART AND XX.NUOVOCODART = X.NUOVOCODART)

UNION

SELECT 
	X.ARTTIPOLOGIA, X.CODART, X.CODICE, '' AS DESCRIMBALLO, X.VARIANTEIMBALLO, X.DESCRVARIANTE, X.NUOVOCODART
	, (SELECT TOP 1 A.DESCRIZIONE FROM ANAGRAFICAARTICOLI A WITH (NOLOCK) WHERE A.CODICE = X.CODART OR A.CODICE = X.ARTTIPOLOGIA) + ' ' + x.DESCRVARIANTE AS DESCRIZIONE
	, X.TIPO 
	, ORDINAMENTO
FROM 
	ZS_GENERAARTICOLI_EXTRA X
WHERE 
	NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST XX WHERE XX.CODART = X.CODART AND XX.NUOVOCODART = X.NUOVOCODART)

	
) CTE
WHERE NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST XX WHERE XX.NUOVOCODART = CTE.NUOVOCODART)
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




SELECT * FROM ZS_VISTA_GENERAARTICOLI
WHERE nuovoCODART not IN ('20067#000000396', 
'20493#000000396', 
'21003#000000396', 
'21005#000000396', 
'21006#000000396', 
'21007#000000396', 
'21011#000000396', 
'21067#000000396', 
'21103#000000396', 
'21105#000000396', 
'21106#000000396', 
'21109#000000396', 
'21493#000000396', 
'22021#000000392', 
'22021#000000390', 
'21002#000000396', 
'21004#000000396', 
'21008#000000396', 
'21009#000000396', 
'21010#000000396', 
'21012#000000396', 
'21107#000000396', 
'21108#000000396', 
'21110#000000396', 
'25104#000000396', 
'25102#000000396', 
'43001#000000370', 
'43004#000000370', 
'43006#000000370', 
'43022#000000370', 
'43058#000000370', 
'43083#000000370', 
'43084#000000370', 
'43105#000000370', 
'43127#000000370', 
'43149#000000370', 
'43185#000000370', 
'43218#000000370', 
'43260#000000370'
)
ORDER BY ARTTIPOLOGIA, ORDINAMENTO, NUOVOCODART


IF OBJECT_ID ('dbo._imballicambiati') IS NOT NULL
	DROP TABLE dbo._imballicambiati
GO

CREATE TABLE dbo._imballicambiati
	(
	CODICE                 VARCHAR (10) NOT NULL,
	DESCRIZIONE            VARCHAR (80),
	VARIANTEIMBALLO        VARCHAR (25),
	VARIANTEIMBALLOF       VARCHAR (25),
	old_variantembiallo    VARCHAR (25),
	old_variantembiallof   VARCHAR (25)
	)
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('206', 'FUSTO 120 KG', '000', NULL, '004', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('255', 'FUSTO PLASTICA ADR 200 KG - 1H1 (220 L)', '011', NULL, '012', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('2551', 'FUSTO PLASTICA ADR 200 KG - 1H1 (220 L) SU PLT 100X120 CM', '011', '', '012', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('2552', 'FUSTO PLASTICA ADR 200 KG - 1H1 (220 L) SU PLT PLASTICA 100X120 CM', '011', '', '012', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('2553', 'FUSTO PLASTICA ADR 200KG - 1H1 (220 L) SU EPAL 80X120 CM', '011', '', '012', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('2554', 'FUSTO PLASTICA ADR 200 KG - 1H1 (220 L) SU PLT 90X110 CM', '011', '', '012', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('504', 'TANICA 25 KG', '303', NULL, '310', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('5041', 'TANICA 25 KG SU PLT 100X120 CM', '303', NULL, '310', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('5042', 'TANICA 25 KG SU PLT PLASTICA 100X120 CM', '303', NULL, '310', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('5043', 'TANICA 25 KG SU EPAL 80X120 CM', '303', NULL, '310', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('5044', 'TANICA 25 KG SU PLT 90X110 CM', '303', NULL, '310', '')
GO

INSERT INTO dbo._imballicambiati (codice, DESCRIZIONE, VARIANTEIMBALLO, VARIANTEIMBALLOF, old_variantembiallo, old_variantembiallof)
VALUES ('505', 'TANICA 30 KG', '303', NULL, '305', '')
GO


SELECT a.CODICE, a.DESCRIZIONE, a.VARESPLICITE FROM 
ANAGRAFICAARTICOLI a JOIN _imballicambiati i ON RIGHT(a.CODICE, 3) = i.VARIANTEIMBALLO
AND len(a.codice) > 12




IF OBJECT_ID ('dbo.ITA_UPDATECODICISOST') IS NOT NULL
	DROP PROCEDURE dbo.ITA_UPDATECODICISOST
GO

CREATE PROCEDURE ITA_UPDATECODICISOST  (@OLDCODART VARCHAR(50), @CODIMBALLO VARCHAR(10), @NEWCODART VARCHAR(50) ) AS
	DECLARE @CSQL_S AS VARCHAR(5000)
	DECLARE @CSQL_U AS VARCHAR(5000)
	DECLARE @CSQL_U0 AS VARCHAR(5000)
	DECLARE @CSQL_U1 AS VARCHAR(5000)
	DECLARE @CSQL_U2 AS VARCHAR(5000)
	DECLARE @TABELLA AS VARCHAR(500)
	
	PRINT @OLDCODART
	PRINT @CODIMBALLO
	PRINT @NEWCODART

	SET @CSQL_U0 = 'UPDATE R SET R.CODART = ''' + @NEWCODART + ''' ' +
	' FROM RIGHEDOCUMENTI R ' +  
	' WHERE R.TIPODOC IN (''OCC'', ''OCE'', ''OCG'') AND R.CODART = ''' + @OLDCODART + ''' AND ((R.RIGACHIUSA = 1 AND R.QTAGESTRES = 0) OR (R.CODIMBALLO  NOT IN (SELECT X.CODICE FROM ZS_GENERAARTICOLI_POST X WHERE X.CODART =  ''' + @OLDCODART + ''')))'
	
	SET @CSQL_U1 = 'UPDATE S SET S.CODART = ''' + @NEWCODART + '''  ' +
	' FROM RIGHEDOCUMENTI R JOIN STORICOMAG S ON S.IDTESTA = R.IDTESTA AND S.RIGADOC = R.IDRIGA' +  
	' WHERE R.TIPODOC IN (''OCC'', ''OCE'', ''OCG'') AND R.CODART = ''' + @OLDCODART + ''' AND ((R.RIGACHIUSA = 1 AND R.QTAGESTRES = 0) OR (R.CODIMBALLO  NOT IN (SELECT X.CODICE FROM ZS_GENERAARTICOLI_POST X WHERE X.CODART =  ''' + @OLDCODART + ''')))'

	
	SET @CSQL_U2 = 'UPDATE P SET P.CODARTICOLO = ''' + @NEWCODART + '''  ' + 
	' FROM RIGHEDOCUMENTI R JOIN STORICOMAG S ON S.IDTESTA = R.IDTESTA AND S.RIGADOC = R.IDRIGA ' +
	' JOIN STORICOPREZZIARTICOLO P ON P.RIFSTORICOMAG = S.PROGRESSIVO' +
	' WHERE R.TIPODOC IN (''OCC'', ''OCE'', ''OCG'') AND R.CODART = ''' + @OLDCODART + ''' AND ((R.RIGACHIUSA = 1 AND R.QTAGESTRES = 0) OR (R.CODIMBALLO  NOT IN (SELECT X.CODICE FROM ZS_GENERAARTICOLI_POST X WHERE X.CODART =  ''' + @OLDCODART + ''')))'
	

	IF LEFT(@oldcodart, 1) IN ('3', '4') AND 
		(
			SELECT count(*) FROM ANAGRAFICAARTICOLI X WHERE X.CODICE = @NEWCODART -- AND TIPO IN ('A', 'B')
		) = 1
	BEGIN
	
		PRINT 'SERIE 30000, 40000aaaaa'
	
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
				
				--PRINT @TABELLA
				IF @TABELLA IN  ('STORICOMAG', 'RIGHEDOCUMENTI', 'DESCRARTICOLI')
				BEGIN
					PRINT 'NON PROCESSATA - ' + @TABELLA + @CSQL_U
					SET @CSQL_U = ''
				END
				
				
				IF @TABELLA = 'TABLOTTIRIORDINO'
				BEGIN
					DELETE FROM TABLOTTIRIORDINO WHERE CODART = @NEWCODART AND EXISTS(SELECT TOP 1 1 FROM TABLOTTIRIORDINO X WHERE X.CODART = @OLDCODART)
				END
				
				PRINT @CSQL_U
				IF (@CSQL_U <> '')
				BEGIN
					EXEC (@CSQL_U)
				END	
	
				
	
			
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
			--WHEN TABELLA = 'GESTIONEPREZZI' THEN
			--	'UPDATE ' + TABELLA + ' SET CODART = ''' + @OLDCODART + '#??????'', CODARTRIC = ''' + @OLDCODART + '#___%'' WHERE INIZIOVALIDITA = FINEVALIDITA AND '+ CAMPO + ' = ''' + @OLDCODART + ''''
			WHEN TABELLA = 'RELAZIONICFV' AND CAMPO = 'VARIANTI' THEN
				'UPDATE ' + TABELLA + ' SET VARIANTI = ''' + SUBSTRING(@NEWCODART, CHARINDEX('#', @NEWCODART) + 1, LEN(@NEWCODART) - CHARINDEX('#', @NEWCODART)) + ''' WHERE ARTICOLO  + ''#'' + ' + CAMPO + ' = ''' + @OLDCODART + ''''
			ELSE
				'UPDATE ' + TABELLA + ' SET ' + CAMPO + ' = ''' + @NEWCODART + ''' WHERE '+ CAMPO + ' = ''' + @OLDCODART + ''''
			END)  
			AS STRSQL
		--, * 
		--, 'UPDATE ITA_CODICISOST SET SEL = 0 WHERE TABELLA = ''' + TABELLA + ''' AND (SELECT COUNT(*) FROM ' + TABELLA + ' WHERE ' + CAMPO + ' IN (SELECT Z.CODART FROM ZS_GENERAARTICOLI z)) = 0' AS STRSQL
		FROM ITA_CODICISOST i 
		WHERE I.SEL = 1
	
	--	ORDER BY C.COLUMN_NAME
		OPEN rSqlA
	
		FETCH NEXT from rSqlA INTO @TABELLA,  @CSQL_S, @CSQL_U
		WHILE (@@FETCH_STATUS <> -1)
			BEGIN
				
				--PRINT @TABELLA
				
				/*
				IF @TABELLA = 'RIGHEDOCUMENTI'
				BEGIN
	
	
	
					PRINT @CSQL_U2
					EXEC (@CSQL_U2)
		
					PRINT @CSQL_U1

					EXEC (@CSQL_U1)
					
					PRINT @CSQL_U0
					EXEC (@CSQL_U0)					
					
				END
				*/
				
				--PRINT @CSQL_U
				EXEC (@CSQL_U)
	
	
				
	
				
				--PRINT @@ROWCOUNT
		
				FETCH NEXT from rSqlA INTO @TABELLA, @CSQL_S, @CSQL_U
			END
	
		CLOSE rSqlA
		DEALLOCATE rSqlA
		
		RETURN
		
	END
GO

GRANT EXECUTE ON dbo.ITA_UPDATECODICISOST TO Metodo98
GO





IF OBJECT_ID ('dbo.ZS_GENERAARTICOLO_POST') IS NOT NULL
	DROP PROCEDURE dbo.ZS_GENERAARTICOLO_POST
GO

create PROCEDURE ZS_GENERAARTICOLO_POST(@CODART VARCHAR(50), @OLDCODART VARCHAR(50), @CODIMBALLO VARCHAR(50), @REDO SMALLINT = 0) AS
BEGIN
 
	DECLARE @ARTPADRE VARCHAR(50)
	DECLARE @ARTMODELLO VARCHAR(50)
	
	SET @ARTPADRE = LEFT(@CODART, CHARINDEX('#' , @CODART) -1)
	SET @ARTMODELLO = @OLDCODART --LEFT(@CODART, LEN(@CODART) -3) + 'XXX'
	

	

	
	IF (NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST WHERE CODART = @OLDCODART AND NUOVOCODART = @CODART AND CODICE = @CODIMBALLO ) OR @REDO = 1)
	BEGIN
	
		PRINT @ARTPADRE
		PRINT @ARTMODELLO
		PRINT @CODIMBALLO
		
	   		
		INSERT INTO dbo.ZS_GENERAARTICOLI_POST (CODART, NUOVOCODART, UtenteModifica, DataModifica, CODICE)
		SELECT @OLDCODART, @CODART, 'TRM', GETDATE(), @CODIMBALLO
		FROM TABDITTE
		WHERE NOT EXISTS(SELECT 1 FROM ZS_GENERAARTICOLI_POST X 
		WHERE -- X.CODART = @OLDCODART AND 
		X.NUOVOCODART = @CODART AND X.CODICE = @CODIMBALLO)
		
		--EXEC ZS_GENERAARTICOLI_GESTIONEPREZZI @CODART, @OLDCODART
		
		
		
		DECLARE @NRPEZZIIMBALLO INT
		

		SELECT @NRPEZZIIMBALLO = G1.QTA_COLLI 
		FROM ITA_VIS_PRZPART_IMBLIST  G1
		WHERE G1.PRIORITA > 1 AND
			(G1.CODART = @CODART OR
			G1.CODART = LEFT(@CODART, LEN(@CODART) -3) + '???' OR
			G1.CODART LIKE @ARTPADRE + '%') --'#??????')
		AND G1.COD_IMBALLO = @CODIMBALLO

		PRINT '-- DATI ANAGRAFICASTANDARD'
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
			
		
		-- articoli originali con imballo qualificato
		IF (LEFT(@CODART , 1) = '2' AND RIGHT(@OLDCODART, 3) <> 'XXX')
		BEGIN
			PRINT '- listini ' + @codart + ' - ' + @ARTMODELLO
			INSERT INTO dbo.LISTINIARTICOLI (CODART, NRLISTINO, UM, PREZZO, PREZZOEURO, UTENTEMODIFICA, DATAMODIFICA, DeltaIncremento, TP_CodConto, TP_ConsPP, TP_PrezzoPart, TP_PrezzoPartEuro, TP_Scorporo, TP_Sconti, TP_QTASCONTO, TP_QTACOEFF, TP_QTAMO, TP_Abbuono, TP_DataCambio, TP_ValoreCambio, DATAVALIDITA, TP_FormulaSct, PREZZOCALC, TP_ABBUONOEURO)
			SELECT @CODART, T.NRLISTINO, UM, PREZZO, PREZZOEURO, 'TRM', GETDATE(), DeltaIncremento, TP_CodConto, TP_ConsPP, TP_PrezzoPart, TP_PrezzoPartEuro, TP_Scorporo, TP_Sconti, TP_QTASCONTO, TP_QTACOEFF, TP_QTAMO, TP_Abbuono, TP_DataCambio, TP_ValoreCambio, DATAVALIDITA, TP_FormulaSct, PREZZOCALC, TP_ABBUONOEURO
			FROM LISTINIARTICOLI t 
			WHERE T.CODART = @OLDCODART
				AND LEFT(@CODART , 1) = '2'
				AND NOT EXISTS (SELECT 1 FROM LISTINIARTICOLI x WHERE @CODART = X.CODART AND x.NRLISTINO = T.NRLISTINO)
		END
		ELSE
		BEGIN		
			PRINT '- listini ' + @codart + ' - ' + @ARTMODELLO
			INSERT INTO dbo.LISTINIARTICOLI (CODART, NRLISTINO, UM, PREZZO, PREZZOEURO, UTENTEMODIFICA, DATAMODIFICA, DeltaIncremento, TP_CodConto, TP_ConsPP, TP_PrezzoPart, TP_PrezzoPartEuro, TP_Scorporo, TP_Sconti, TP_QTASCONTO, TP_QTACOEFF, TP_QTAMO, TP_Abbuono, TP_DataCambio, TP_ValoreCambio, DATAVALIDITA, TP_FormulaSct, PREZZOCALC, TP_ABBUONOEURO)
			SELECT @CODART, T.NRLISTINO, UM, PREZZO, PREZZOEURO, 'TRM', GETDATE(), DeltaIncremento, TP_CodConto, TP_ConsPP, TP_PrezzoPart, TP_PrezzoPartEuro, TP_Scorporo, TP_Sconti, TP_QTASCONTO, TP_QTACOEFF, TP_QTAMO, TP_Abbuono, TP_DataCambio, TP_ValoreCambio, DATAVALIDITA, TP_FormulaSct, PREZZOCALC, TP_ABBUONOEURO
			FROM LISTINIARTICOLI t 
			WHERE T.CODART = @ARTMODELLO
				AND LEFT(@CODART , 1) = '2'
				AND NOT EXISTS (SELECT 1 FROM LISTINIARTICOLI x WHERE @CODART = X.CODART AND x.NRLISTINO = T.NRLISTINO)
		END	
			
		
		DELETE A1 
		FROM CONTROPARTARTICOLI A1 
		WHERE a1.CODART = @CODART 
		
		INSERT INTO dbo.CONTROPARTARTICOLI (CODART, ESERCIZIO, NUMERO, SCGEN, UTENTEMODIFICA, DATAMODIFICA)
		SELECT @CODART, T.ESERCIZIO, T.NUMERO, T.SCGEN, 'TRM', GETDATE()
		FROM CONTROPARTARTICOLI t 
		WHERE T.CODART = @ARTMODELLO
			AND LEFT(@CODART , 1) = '2'



		IF LEFT(@CODART , 1) IN ('2', '3', '4') AND RIGHT(@CODART, 3) = 'XXX'
		BEGIN
			
			PRINT '-- AGGIORNO CODICI STANDARD ' + @CODART
		
			EXEC ITA_UPDATECODICISOST  @OLDCODART, @CODIMBALLO, @CODART    
			
			
			UPDATE a1
			SET 
				a1.NRPEZZIIMBALLO = @NRPEZZIIMBALLO --a2.NRPEZZIIMBALLO
				, a1.RIFERIMIMBALLO = LEFT(@CODIMBALLO, 3) --(SELECT TOP 1 tabimballi.CODICE FROM TABIMBALLI WHERE tabimballi.VARIANTEIMBALLO = RIGHT(a1.CODICE, 3))
			FROM ANAGRAFICAARTICOLI a1
			WHERE 
				a1.CODICE = @CODART
				AND LEFT(a1.CODICE , 1) IN ('2', '3', '4')	 
			
			UPDATE EXTRAMAG
			SET NRPEZZIIMBALLO = @NRPEZZIIMBALLO	
			WHERE CODART = @CODART
				
		END
						
		-- inserimento in tabella articoli adr
		INSERT INTO dbo.TAB_ARTICOLIADR (CODART, CLASSEADR, NUMORD, NUMONU, NUMPER, GI, UTENTEMODIFICA, DATAMODIFICA, DESIGN_TRASP, COD_GALLERIE, ESENZIONE_QTA)
		SELECT @CODART, CLASSEADR, NUMORD, NUMONU, NUMPER, GI, 'trm', getdate(), DESIGN_TRASP, COD_GALLERIE, ESENZIONE_QTA
		FROM TAB_ARTICOLIADR t 
		WHERE (t.CODART = @ARTMODELLO OR t.CODART = @OLDCODART)
		AND @CODART NOT IN (SELECT x.codart FROM TAB_ARTICOLIADR x)	  
		
	/*
		-- AGGIORNAMENTO STORICOMAG
		--SELECT S.CODART, S.RIFERIMENTI, t.DESCRIZIONE, a.DESCRIZIONE, r.IDTESTA, r.IDRIGA
		
		UPDATE s SET s.CODART = @CODART
		FROM RIGHEDOCUMENTI r JOIN STORICOMAG S ON S.IDTESTA = R.IDTESTA AND S.RIGADOC = R.IDRIGA
		JOIN tabimballi t ON t.CODICE = r.CODIMBALLO
		WHERE r.RIGACHIUSA = 0 AND r.QTAGESTRES > 0 AND r.TIPODOC IN ('OCG', 'OCC', 'OCE')
		AND r.CODART = @OLDCODART AND r.CODIMBALLO = @CODIMBALLO

		-- AGGIORNAMENTO RIGHEDOCUMENTI
		--SELECT r.CODART, r.CODIMBALLO, e.NUOVOCODART, t.DESCRIZIONE, a.DESCRIZIONE, r.IDTESTA, r.IDRIGA		
		
		UPDATE r
		SET r.CODART = @CODART
		FROM RIGHEDOCUMENTI r 
		JOIN tabimballi t ON t.CODICE = r.CODIMBALLO
		WHERE r.RIGACHIUSA = 0 AND r.QTAGESTRES > 0 AND r.TIPODOC IN ('OCG', 'OCC', 'OCE')
		AND r.CODART = @OLDCODART AND r.CODIMBALLO = @CODIMBALLO
	 */	
		
		
		
	END
END
GO

GRANT EXECUTE ON dbo.ZS_GENERAARTICOLO_POST TO Metodo98
GO









