

IF OBJECT_ID ('dbo.VTECRM_EXTRACLIENTI') IS NOT NULL
	DROP VIEW dbo.VTECRM_EXTRACLIENTI
GO

create view VTECRM_EXTRACLIENTI as 
	
SELECT *
 FROM OPENQUERY (VTECRMTEST ,' SELECT
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
	SELECT *
 FROM OPENQUERY (VTECRMTEST ,' SELECT
prod.external_code,
prod.productname,
  cf.*
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

