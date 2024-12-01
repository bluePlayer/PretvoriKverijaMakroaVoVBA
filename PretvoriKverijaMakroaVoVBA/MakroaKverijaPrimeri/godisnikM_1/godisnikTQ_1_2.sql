INSERT INTO Tgodisnik_1 ( GOD, sort, NAZIV, kol1, kol2, kol3, kol4, kol5, kol6, kol7, kol8, kol9 )
SELECT TabelaVnes.GOD, "0" & [GODISNIK] AS sort, TabelaAdresar1.GODISNIK_NAZIV AS NAZIV, Count(TabelaVnes.POL) AS kol1, Sum(IIf([pol]="1",1,0)) AS kol2, Sum(IIf([pol]="2",1,0)) AS kol3, Sum(IIf([TabelaVnes].[nacin]="1",1,0)) AS kol4, Sum(IIf([TabelaVnes].[nacin]="1" And [pol]="1",1,0)) AS kol5, Sum(IIf([TabelaVnes].[nacin]="1" And [pol]="2",1,0)) AS kol6, Sum(IIf([TabelaVnes].[nacin]="2",1,0)) AS kol7, Sum(IIf([TabelaVnes].[nacin]="2" And [pol]="1",1,0)) AS kol8, Sum(IIf([TabelaVnes].[nacin]="2" And [pol]="2",1,0)) AS kol9
FROM TabelaVnes INNER JOIN TabelaAdresar1 ON (TabelaVnes.OTSEK = TabelaAdresar1.SIFRAOTSEK) AND (TabelaVnes.EVS = TabelaAdresar1.EDINICASOS) AND (TabelaVnes.MATBR = TabelaAdresar1.MATBR)
GROUP BY TabelaVnes.GOD, "0" & [GODISNIK], TabelaAdresar1.GODISNIK_NAZIV;

