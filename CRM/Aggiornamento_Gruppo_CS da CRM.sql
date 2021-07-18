-- 02/02/2017 - aggiornamento gruppi e categoriestatistiche da file excel crm

--SELECT * FROM TABGRUPPI


INSERT INTO dbo.TABGRUPPI (CODICE, DESCRIZIONE, MAGFISCALE, NOTE, UTENTEMODIFICA, DATAMODIFICA, ScontiPremi, SpTrasp, Provvi, Imballi, Amministrazione, SpVendita, SpMagazzino, SpeseProd, CostoAggiunto)
VALUES (101, 'DISTRIBUTORI', 1, '', 'tsmonica', '02/02/2017 11:01:22', 0, 0, 0, 0, 0, 0, 0, 0, 0)
GO

INSERT INTO dbo.TABGRUPPI (CODICE, DESCRIZIONE, MAGFISCALE, NOTE, UTENTEMODIFICA, DATAMODIFICA, ScontiPremi, SpTrasp, Provvi, Imballi, Amministrazione, SpVendita, SpMagazzino, SpeseProd, CostoAggiunto)
VALUES (102, 'PRODUZIONE PROPRIA', 1, '', 'tsmonica', '02/02/2017 11:01:22', 0, 0, 0, 0, 0, 0, 0, 0, 0)
GO

INSERT INTO dbo.TABGRUPPI (CODICE, DESCRIZIONE, MAGFISCALE, NOTE, UTENTEMODIFICA, DATAMODIFICA, ScontiPremi, SpTrasp, Provvi, Imballi, Amministrazione, SpVendita, SpMagazzino, SpeseProd, CostoAggiunto)
VALUES (103, 'RIVENDITA', 1, '', 'tsmonica', '02/02/2017 11:01:22', 0, 0, 0, 0, 0, 0, 0, 0, 0)
GO

--SELECT * FROM TABCATEGORIESTAT

INSERT INTO dbo.TABCATEGORIESTAT (CODICE, DESCRIZIONE, NOTE, UTENTEMODIFICA, DATAMODIFICA, Budget)
VALUES (101, 'ATTIVI', '', 'tsmonica', '02/02/2017 11:03:48', 0)
GO

INSERT INTO dbo.TABCATEGORIESTAT (CODICE, DESCRIZIONE, NOTE, UTENTEMODIFICA, DATAMODIFICA, Budget)
VALUES (102, 'DERIVATO ACIDO LATTICO', '', 'tsmonica', '02/02/2017 11:05:25', 0)
GO

INSERT INTO dbo.TABCATEGORIESTAT (CODICE, DESCRIZIONE, NOTE, UTENTEMODIFICA, DATAMODIFICA, Budget)
VALUES (103, 'FOSFONATO', '', 'tsmonica', '02/02/2017 11:05:56', 0)
GO

INSERT INTO dbo.TABCATEGORIESTAT (CODICE, DESCRIZIONE, NOTE, UTENTEMODIFICA, DATAMODIFICA, Budget)
VALUES (104, 'OCCASIONALI', '', 'tsmonica', '02/02/2017 11:07:21', 0)
GO

INSERT INTO dbo.TABCATEGORIESTAT (CODICE, DESCRIZIONE, NOTE, UTENTEMODIFICA, DATAMODIFICA, Budget)
VALUES (105, 'PROTEINA', '', 'tsmonica', '02/02/2017 11:07:48', 0)
GO





-- SELECT * INTO _aaprecrmgec FROM ANAGRAFICAARTICOLI


UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=105 WHERE CODICE = '11381'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=104 WHERE CODICE = '32001'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=104 WHERE CODICE = '32650'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11205'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11440'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43022'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11060'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11040'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11045'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11159'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11037'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11075'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43019'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43003'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43013'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43000'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43018'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43001'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43006'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43004'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '20037#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11255'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43056'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32314'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32307'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32311'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32309'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43150'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32310'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32331'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32332'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32339'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32369'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43381'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32410'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32417'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32419'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32421'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32422'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32423'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43094'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32464'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32462'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32467'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32466'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43095'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '20265#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '20265#251XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '20265#252XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '30266'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '20269#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '20268#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '20271#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43169'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43164'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43166'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43245'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43224'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43225'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11019'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43086'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43083'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43074'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43093'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43375'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43082'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43085'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43067'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43073'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43091'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=10 WHERE CODICE = '43168'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43151'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43079'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11529'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11530'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32678'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '11277'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '20556#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '20564#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32680'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '32682'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43396'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=10 WHERE CODICE = '43395'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '22034#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '22035#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '22036#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '22037#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=10 WHERE CODICE = '22038#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20025#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20025#231XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=5 WHERE CODICE = '43007'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20253#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32035'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20183#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20210#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20214#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20211#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20211#005XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20213#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20289#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20193#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20255#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20229#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20286#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32268'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20278#095XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20278#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20233#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20274#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32670'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20254#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32445'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32266'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32261'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32263'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20256#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20279#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20219#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20231#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32260'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20295#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20216#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20218#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '43105'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20030#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '11299'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20181#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20182#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32442'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32436'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20258#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20296#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20222#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20275#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20266#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20284#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=5 WHERE CODICE = '43185'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20415#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32675'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=5 WHERE CODICE = '32662'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20523#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20523#112XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20523#104XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20523#082XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20523#086XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20523#005XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20523#072XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20526#016XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '20532#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '22052#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '22053#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '22054#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=5 WHERE CODICE = '22055#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=103 WHERE CODICE = '43049'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21002#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21004#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21008#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=28 WHERE CODICE = '43100'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=28 WHERE CODICE = '32506'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=28 WHERE CODICE = '43180'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=28 WHERE CODICE = '43193'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=28 WHERE CODICE = '43170'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=28 WHERE CODICE = '43172'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=28 WHERE CODICE = '43171'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=28 WHERE CODICE = '43191'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21106#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21110#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '20493#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21109#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '20067#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21103#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21105#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21102#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21104#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21107#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=28 WHERE CODICE = '11262'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=28 WHERE CODICE = '43041'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=28 WHERE CODICE = '43078'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=28 WHERE CODICE = '43058'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=28 WHERE CODICE = '43175'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=28 WHERE CODICE = '43178'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=28 WHERE CODICE = '43149'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=28 WHERE CODICE = '43084'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=28 WHERE CODICE = '43110'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21006#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21493#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21003#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21005#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21007#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=28 WHERE CODICE = '21067#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=102 WHERE CODICE = '11072'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=102 WHERE CODICE = '43121'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=102 WHERE CODICE = '43122'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=102 WHERE CODICE = '43125'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=102 WHERE CODICE = '43123'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=102 WHERE CODICE = '43124'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=102 WHERE CODICE = '43126'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=102 WHERE CODICE = '43127'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=102 WHERE CODICE = '88079'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=13 WHERE CODICE = '43207'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=13 WHERE CODICE = '43208'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=13 WHERE CODICE = '20292#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=13 WHERE CODICE = '43070'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=13 WHERE CODICE = '22039#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=13 WHERE CODICE = '22040#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=101 WHERE CODICE = '32250'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=101 WHERE CODICE = '11281'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=101 WHERE CODICE = '11310'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '43035'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20194#272XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20194#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20195#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20195#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20196#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20196#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20198#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '43065'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '32435'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '32439'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=9 WHERE CODICE = '43161'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20235#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20237#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20237#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20240#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20053#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20053#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20249#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20249K#000'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20249#100XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20247#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20247#100XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '20263#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20285#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '43143'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '43141'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20281#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20208#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20208#095XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20207#095XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20207#241XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20236#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20236#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20260#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20260#242XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '32505'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20300#104XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20300#264XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20300#265XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '32504'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '32560'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '32561'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=9 WHERE CODICE = '43230'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=9 WHERE CODICE = '43221'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=9 WHERE CODICE = '43223'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=9 WHERE CODICE = '43218'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=9 WHERE CODICE = '43220'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '32500'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20351#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20385#114XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20355#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20355#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20355#015XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20355#112XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20355#104XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20357#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20365#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20367#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20365#264XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20385#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20385#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20391#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20385#095XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20385#104XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20385#096XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20385#106XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20384#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20395#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20395#107XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20395K#107'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20395#106XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20410#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20410#108XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20410#104XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '32562'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20360#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '43075'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=9 WHERE CODICE = '32565'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20561#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20558#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20535#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20534#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20535#163XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20535#095XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20527#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20550#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20549#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20566#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20552#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20544#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20553#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20553#172XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '25553#172XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20553#095XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20553#170XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20547#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20546#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20554#115XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20554#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20559#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20537#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20503#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20489#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20498#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20031#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20498#112XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20495#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20494#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20492#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20505#010XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20510#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20512#007XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20512#095XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20488#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20510#026XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20542#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20491#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20516#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20505#104XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20507#132XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '25507#132XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20505#013XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20505#018XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20505#019XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20505#005XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20509#151XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20543#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20517#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '25517#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20517#019XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20517#020XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20563#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20508#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20508#104XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20508#013XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20508#005XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20496#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20496#013XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20541#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20520#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20520#062XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20497#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20497#104XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20497#013XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20533#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20533#095XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20533#117XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20519#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20522#071XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20540#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20522#112XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20522#113XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20525#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20525#072XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20522#104XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20522#005XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20522#072XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20523#113XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20523#116XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '20500#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22009#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22010#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22011#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22013#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22014#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22015#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22016#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22017#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22018#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22019#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22022#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22023#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22024#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22025#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22026#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22027#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22028#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22029#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22030#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22031#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=9 WHERE CODICE = '22032#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =101, CODCATEGORIASTAT=11 WHERE CODICE = '43025'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20069#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20018#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20068#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20020#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20021#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20021#214XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20021#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20021#218XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20021#220XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20021#112XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20033#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20051#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20051#217XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20071#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=11 WHERE CODICE = '32012'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20043#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20063#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=11 WHERE CODICE = '32005'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20060#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=11 WHERE CODICE = '32020'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=11 WHERE CODICE = '11082'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =103, CODCATEGORIASTAT=11 WHERE CODICE = '43053'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20046#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20046#243XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '20066#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '22001#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '22002#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '22003#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '22004#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '22005#000XXX'
UPDATE ANAGRAFICAARTICOLI SET GRUPPO =102, CODCATEGORIASTAT=11 WHERE CODICE = '22006#000XXX'
