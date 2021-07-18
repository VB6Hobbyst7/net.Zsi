
SELECT *, (SELECT TOP 1 A.DESCRIZIONE FROM ANAGRAFICAARTICOLI A WHERE A.CODICE = CTE.CODART OR A.CODICE = CTE.ARTTIPOLOGIA) AS DESCRIZIONE FROM (
	SELECT distinct
	(CASE WHEN charindex('#', g.CODART) > 1 THEN LEFT(g.CODART, charindex('#', g.CODART) - 1) ELSE g.CODART END) AS ARTTIPOLOGIA
	, g.CODART
	, t.CODICE
	, t.varianteimballo
	, v.DESCRIZIONE AS descrizionevariante
	, v.a
	, v.b
	, (CASE 
		WHEN charindex('#', g.CODART) > 0 AND charindex('?', g.codart) = 0 THEN LEFT(g.CODART, charindex('#', g.CODART) + 3) + T.VARIANTEIMBALLO 
		WHEN charindex('#', g.CODART) > 0 AND charindex('?', g.codart) > 0 THEN LEFT(g.CODART, charindex('#', g.CODART)) + '000' + T.VARIANTEIMBALLO 
		ELSE
			g.CODART + '#000' + T.VARIANTEIMBALLO 
		END) AS NUOVOCODART
		--, g.*
		, t.DESCRIZIONE
	FROM GESTIONEPREZZI g JOIN GESTIONEPREZZIRIGHE r ON g.PROGRESSIVO = r.RIFPROGRESSIVO
	JOIN tabimballi t ON t.CODICE = r.cod_imballo 
	JOIN TABVARIANTI v ON v.VARIANTE = t.varianteimballo
	WHERE g.INIZIOVALIDITA = g.FINEVALIDITA and t.varianteimballo  <> '' --AND charindex('?', g.codart) = 0
	AND v.TIPOLOGIA ='62'
	--AND t.codice NOT IN (100, 101) AND 
	) CTE
WHERE 
	NUOVOCODART NOT IN (SELECT CODICE FROM ANAGRAFICAARTICOLI)
	AND ARTTIPOLOGIA BETWEEN '20000' AND '49999'
