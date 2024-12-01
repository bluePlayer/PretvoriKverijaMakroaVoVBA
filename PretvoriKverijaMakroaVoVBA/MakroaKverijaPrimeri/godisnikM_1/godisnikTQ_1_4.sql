INSERT INTO Tgodisnik_1 ( GOD, sort, NAZIV, kol1, kol2, kol3, kol4, kol5, kol6, kol7, kol8, kol9 )
SELECT TabelaVnes.GOD, [GODISNIK] & [ID1] AS sort, TabelaAdresar.VIDSKOLA AS NAZIV, Count(TabelaVnes.POL) AS kol1, Sum(IIf([pol]="1",1,0)) AS kol2, Sum(IIf([pol]="2",1,0)) AS kol3, Sum(IIf([TabelaVnes].[nacin]="1",1,0)) AS kol4, Sum(IIf([TabelaVnes].[nacin]="1" And [pol]="1",1,0)) AS kol5, Sum(IIf([TabelaVnes].[nacin]="1" And [pol]="2",1,0)) AS kol6, Sum(IIf([TabelaVnes].[nacin]="2",1,0)) AS kol7, Sum(IIf([TabelaVnes].[nacin]="2" And [pol]="1",1,0)) AS kol8, Sum(IIf([TabelaVnes].[nacin]="2" And [pol]="2",1,0)) AS kol9
FROM TabelaVnes INNER JOIN TabelaAdresar ON (TabelaVnes.OTSEK = TabelaAdresar.SIFRAOTSEK) AND (TabelaVnes.EVS = TabelaAdresar.EDINICASOS) AND (TabelaVnes.MATBR = TabelaAdresar.MATBR)
WHERE (((TabelaAdresar.GODISNIK)<>"3"))
GROUP BY TabelaVnes.GOD, [GODISNIK] & [ID1], TabelaAdresar.VIDSKOLA;

