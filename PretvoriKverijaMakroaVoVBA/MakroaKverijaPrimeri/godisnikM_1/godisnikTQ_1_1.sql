SELECT TabelaVnes.GOD, "00" AS sort, "VKUPNO" AS NAZIV, "" AS REDBRFAK, "" AS MATBR, "" AS EVS, Count(TabelaVnes.POL) AS kol1, Sum(IIf([pol]="1",1,0)) AS kol2, Sum(IIf([pol]="2",1,0)) AS kol3, Sum(IIf([TabelaVnes].[metod]="1",1,0)) AS kol4, Sum(IIf([TabelaVnes].[metod]="1" And [pol]="1",1,0)) AS kol5, Sum(IIf([TabelaVnes].[metod]="1" And [pol]="2",1,0)) AS kol6, Sum(IIf([TabelaVnes].[metod]="2",1,0)) AS kol7, Sum(IIf([TabelaVnes].[metod]="2" And [pol]="1",1,0)) AS kol8, Sum(IIf([TabelaVnes].[metod]="2" And [pol]="2",1,0)) AS kol9 INTO Tgodisnik_1
FROM TabelaVnes INNER JOIN TabelaAdresar ON (TabelaVnes.EVS = TabelaAdresar.EVS) AND (TabelaVnes.MATBR = TabelaAdresar.MATBR)
GROUP BY TabelaVnes.GOD, "00", "VKUPNO", "", "", "";

