
IF OBJECT_ID ('dbo.VTECRM_EXTRACLIENTI') IS NOT NULL
	DROP VIEW dbo.VTECRM_EXTRACLIENTI
GO

create view VTECRM_EXTRACLIENTI as 
	
SELECT external_code  COLLATE Latin1_General_CI_AS AS CODCONTO
, accountname
, cf_922 AS SettoreCRM
, cf_924 AS FunzionarioCRM
, cf_1071 AS GruppoCRM
, CAST(isnull(cf_926, '-') AS VARCHAR(500)) AS CategoriaCRM

 FROM OPENQUERY (VTECRMPROD ,' SELECT
	acc.external_code,
	acc.accountname,
	cf.*
FROM
  vte_account acc
  INNER JOIN vte_crmentity 
    ON crmid = acc.accountid
    INNER JOIN vte_accountscf cf 
    ON acc.accountid = cf.accountid
WHERE deleted = 0 -- condizione che esclude gli eliminati 
' )
GO

GRANT SELECT ON dbo.VTECRM_EXTRACLIENTI TO Metodo98
GO





IF OBJECT_ID ('dbo.VTECRM_EXTRAARTICOLI') IS NOT NULL
	DROP VIEW dbo.VTECRM_EXTRAARTICOLI
GO

create view VTECRM_EXTRAARTICOLI as 
	SELECT external_code COLLATE Latin1_General_CI_AS AS CODICE
	, gruppo AS GruppoCRM
	, natura AS NaturaCRM
	, categoria_istat AS CategoriaIstatCRM
	, famiglia AS FamigliaCRM
	, concentrazione AS ConcentrazioneCRM
	, tipologia AS TipologiaCRM
	, provenienza AS ProvenienzaCRM
 FROM OPENQUERY (VTECRMPROD ,' SELECT
	prod.external_code,
	prod.productname,
	prod.gruppo,
	prod.natura,
	prod.categoria_istat,
	prod.famiglia,
	prod.concentrazione,
	prod.tipologia,
	prod.provenienza
FROM
  vte_products prod
  INNER JOIN vte_crmentity 
    ON crmid = prod.productid
  INNER JOIN vte_productcf cf 
    ON prod.productid = cf.productid
WHERE deleted = 0
' )
GO

GRANT SELECT ON dbo.VTECRM_EXTRAARTICOLI TO Metodo98
GO

