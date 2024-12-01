SELECT TabelaVnes.GOD, "00" AS sort, "VKUPNO" AS NAZIV, "" AS REDBRFAK, "" AS MATBR, "" AS EDINICASOS, Count(TabelaVnes.POL) AS kol1, Sum(IIf([pol]="1",1,0)) AS kol2, Sum(IIf([pol]="2",1,0)) AS kol3, Sum(IIf([TabelaVnes].[nacin]="1",1,0)) AS kol4, Sum(IIf([TabelaVnes].[nacin]="1" And [pol]="1",1,0)) AS kol5, Sum(IIf([TabelaVnes].[nacin]="1" And [pol]="2",1,0)) AS kol6, Sum(IIf([TabelaVnes].[nacin]="2",1,0)) AS kol7, Sum(IIf([TabelaVnes].[nacin]="2" And [pol]="1",1,0)) AS kol8, Sum(IIf([TabelaVnes].[nacin]="2" And [pol]="2",1,0)) AS kol9 INTO Tgodisnik_10
FROM TabelaVnes INNER JOIN TabelaAdresar ON (TabelaVnes.EVS = TabelaAdresar.EDINICASOS) AND (TabelaVnes.MATBR = TabelaAdresar.MATBR) AND (TabelaVnes.OTSEK = TabelaAdresar.SIFRAOTSEK)
WHERE (((TabelaVnes.DRZAVJ)="807"))
GROUP BY TabelaVnes.GOD, "00", "VKUPNO", "", "", "";

