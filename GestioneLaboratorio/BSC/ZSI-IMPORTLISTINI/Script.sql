IF OBJECT_ID ('dbo.VISTA_GESTIONEPREZZI') IS NOT NULL
	DROP VIEW dbo.VISTA_GESTIONEPREZZI
GO

CREATE VIEW VISTA_GESTIONEPREZZI AS
        SELECT GP.PROGRESSIVO,
	GPR.IDRIGA,
        GP.CODGRUPPOPREZZICF,
        GP.CODCLIFOR,
        GP.CODART,
        GP.CODGRUPPOPREZZIMAG,GPR.QTAMINIMA,GPR.UM,
        GP.INIZIOVALIDITA,
        GP.FINEVALIDITA,GP.USANRLISTINO,
        GP.TIPOARROT,
        GP.ARROTALIRE,
        GP.ARROTAEURO,
        GPR.PREZZO_MAGG,
        GPR.PREZZO_MAGGEURO,
        GPR.TIPO,
        GPR.SCONTO_UNICO,
        GPR.SCONTO_AGGIUNTIVO,
        GPR.NRLISTINO,
        L.DESCRIZIONE AS DSCLISTINO,
        GP.CODARTRIC,
        GP.PROGRESSIVOCTR,
        (CASE WHEN CHARINDEX('#',GP.CODART)>0 THEN SUBSTRING(GP.CODART,1,CHARINDEX('#',GP.CODART)-1) ELSE GP.CODART END) AS CODPADRE,
        (SELECT DESCRIZIONE FROM ANAGRAFICAARTICOLI WHERE CODICE=GP.CODART) AS DSCA,
        (CASE WHEN (SELECT DESCRIZIONE FROM ANAGRAFICAARTICOLI WHERE CODICE=GP.CODART) IS NULL THEN (SELECT DESCRIZIONE FROM ANAGRAFICAARTICOLI WHERE CODICE=(CASE WHEN CHARINDEX('#',GP.CODART)>0 THEN SUBSTRING(GP.CODART,1,CHARINDEX('#',GP.CODART)-1) ELSE GP.CODART END)) ELSE(SELECT DESCRIZIONE FROM ANAGRAFICAARTICOLI WHERE CODICE=GP.CODART) END) AS DSCARTICOLO,
        (SELECT DSCCONTO1 FROM ANAGRAFICACF WHERE CODCONTO=GP.CODCLIFOR) AS RAGSOCCF,
        (CASE WHEN (SELECT DSCCONTO1 FROM ANAGRAFICACF WHERE CODCONTO=GP.CODCLIFOR) IS NULL THEN (CASE WHEN GP.CODCLIFOR='' OR GP.CODCLIFOR IS NULL THEN 'TUTTI I CLIENTI E FORNITORI' ELSE (CASE WHEN GP.CODCLIFOR='C' THEN 'TUTTI I CLIENTI' ELSE 'TUTTI I FORNITORI' END) END) ELSE (SELECT DSCCONTO1 FROM ANAGRAFICACF WHERE CODCONTO=GP.CODCLIFOR) END) AS DSCCONTO,
        (SELECT DESCRIZIONE FROM TABRAGGRPREZZICF WHERE CODICE=GP.CODGRUPPOPREZZICF) AS DSCCATEGORIA,
        (SELECT DESCRIZIONE FROM TABRAGGRUPPAPREZZI WHERE CODICE=GP.CODGRUPPOPREZZIMAG) AS DSCGRUPPO
        , GPR.COD_IMBALLO
        ,  (SELECT DESCRIZIONE FROM TABIMBALLI I WHERE I.CODICE=GPR.COD_IMBALLO) AS DSCIMBALLO
        FROM GESTIONEPREZZI GP, GESTIONEPREZZIRIGHE GPR LEFT OUTER JOIN TABLISTINI L ON L.NRLISTINO = GPR.NRLISTINO
        WHERE (GP.PROGRESSIVO=GPR.RIFPROGRESSIVO)
        

GO

GRANT DELETE ON dbo.VISTA_GESTIONEPREZZI TO Metodo98
GO
GRANT INSERT ON dbo.VISTA_GESTIONEPREZZI TO Metodo98
GO
GRANT REFERENCES ON dbo.VISTA_GESTIONEPREZZI TO Metodo98
GO
GRANT SELECT ON dbo.VISTA_GESTIONEPREZZI TO Metodo98
GO
GRANT UPDATE ON dbo.VISTA_GESTIONEPREZZI TO Metodo98
GO

