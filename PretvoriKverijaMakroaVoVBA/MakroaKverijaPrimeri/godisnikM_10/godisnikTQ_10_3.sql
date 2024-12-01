INSERT INTO Tgodisnik_10 ( GOD, sort, NAZIV )
SELECT TabelaVnes.GOD, [GODISNIK] & "1" AS sort, TabelaAdresar.GODISNIK_NAZIV AS NAZIV
FROM TabelaVnes INNER JOIN TabelaAdresar ON (TabelaVnes.EVS = TabelaAdresar.EDINICASOS) AND (TabelaVnes.MATBR = TabelaAdresar.MATBR) AND (TabelaVnes.OTSEK = TabelaAdresar.SIFRAOTSEK)
WHERE (((TabelaVnes.DRZAVJ)="807"))
GROUP BY TabelaVnes.GOD, [GODISNIK] & "1", TabelaAdresar.GODISNIK_NAZIV;

