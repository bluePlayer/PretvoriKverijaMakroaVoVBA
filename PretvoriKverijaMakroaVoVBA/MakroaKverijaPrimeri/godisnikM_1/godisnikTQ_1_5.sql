INSERT INTO Tgodisnik_1 ( GOD, sort, NAZIV, RBR, MATBR, EVS, kol1, kol2, kol3, kol4, kol5, kol6, kol7, kol8, kol9 )
SELECT TabelaVnes.GOD, [GODISNIK] & [IDX] AS sort, TabelaAdresar.NAZIV, TabelaAdresar.RBR, TabelaAdresar.MATBR, TabelaAdresar.EVS, Count(TabelaVnes.POL) AS kol1, Sum(IIf([pol]="1",1,0)) AS kol2, Sum(IIf([pol]="2",1,0)) AS kol3, Sum(IIf([TabelaVnes].[metod]="1",1,0)) AS kol4, Sum(IIf([TabelaVnes].[metod]="1" And [pol]="1",1,0)) AS kol5, Sum(IIf([TabelaVnes].[metod]="1" And [pol]="2",1,0)) AS kol6, Sum(IIf([TabelaVnes].[metod]="2",1,0)) AS kol7, Sum(IIf([TabelaVnes].[metod]="2" And [pol]="1",1,0)) AS kol8, Sum(IIf([TabelaVnes].[metod]="2" And [pol]="2",1,0)) AS kol9
FROM TabelaVnes INNER JOIN TabelaAdresar ON (TabelaVnes.EVS = TabelaAdresar.EVS) AND (TabelaVnes.MATBR = TabelaAdresar.MATBR)
GROUP BY TabelaVnes.GOD, [GODISNIK] & [IDX], TabelaAdresar.NAZIV, TabelaAdresar.RBR, TabelaAdresar.MATBR, TabelaAdresar.EVS;

