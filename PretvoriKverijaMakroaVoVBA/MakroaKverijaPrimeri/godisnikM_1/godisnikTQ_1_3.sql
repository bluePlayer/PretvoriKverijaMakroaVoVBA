INSERT INTO Tgodisnik_1 ( GOD, sort, NAZIV )
SELECT TabelaVnes2.GOD, [GODISNIK] & "1" AS sort, TabelaAdresar.GODISNIK_NAZIV AS NAZIV
FROM TabelaVnes2 INNER JOIN TabelaAdresar ON (TabelaVnes2.EVS = TabelaAdresar.EVS) AND (TabelaVnes2.MATBR = TabelaAdresar.MATBR)
GROUP BY TabelaVnes2.GOD, [GODISNIK] & "1", TabelaAdresar.GODISNIK_NAZIV;

