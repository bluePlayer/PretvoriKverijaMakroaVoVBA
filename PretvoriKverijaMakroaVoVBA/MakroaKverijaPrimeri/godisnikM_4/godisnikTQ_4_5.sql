INSERT INTO Tgodisnik_4 ( GOD, sort, NAZIV )
SELECT TabelaVnes.GOD, [GODISNIK] & "1" AS sort, TabelaAdresar.GODISNIK_NAZIV AS NAZIV
FROM TabelaVnes INNER JOIN TabelaAdresar ON (TabelaVnes.OTSEK = TabelaAdresar.SIFRAOTSEK) AND (TabelaVnes.EVS = TabelaAdresar.EDINICASOS) AND (TabelaVnes.MATBR = TabelaAdresar.MATBR)
GROUP BY TabelaVnes.GOD, [GODISNIK] & "1", TabelaAdresar.GODISNIK_NAZIV;

