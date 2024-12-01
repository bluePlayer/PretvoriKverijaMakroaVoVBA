INSERT INTO Tgodisnik_10 ( GOD, sort, NAZIV, REDBRFAK, MATBR, EDINICASOS, kol1, kol2, kol3, kol4, kol5, kol6, kol7, kol8, kol9 )
SELECT TabelaVnes.GOD, [GODISNIK] & [ID1_GOD] AS sort, TabelaVnes.NAZIV, TabelaVnes.REDBRFAK, TabelaVnes.MATBR, TabelaVnes.EDINICASOS, Count(TabelaVnes.POL) AS kol1, Sum(IIf([pol]="1",1,0)) AS kol2, Sum(IIf([pol]="2",1,0)) AS kol3, Sum(IIf([TabelaVnes].[nacin]="1",1,0)) AS kol4, Sum(IIf([TabelaVnes].[nacin]="1" And [pol]="1",1,0)) AS kol5, Sum(IIf([TabelaVnes].[nacin]="1" And [pol]="2",1,0)) AS kol6, Sum(IIf([TabelaVnes].[nacin]="2",1,0)) AS kol7, Sum(IIf([TabelaVnes].[nacin]="2" And [pol]="1",1,0)) AS kol8, Sum(IIf([TabelaVnes].[nacin]="2" And [pol]="2",1,0)) AS kol9
FROM TabelaVnes INNER JOIN TabelaVnes ON (TabelaVnes.EVS = TabelaVnes.EDINICASOS) AND (TabelaVnes.MATBR = TabelaVnes.MATBR) AND (TabelaVnes.OTSEK = TabelaVnes.SIFRAOTSEK)
WHERE (((TabelaVnes.DRZAVJ)="807"))
GROUP BY TabelaVnes.GOD, [GODISNIK] & [ID1_GOD], TabelaVnes.NAZIV, TabelaVnes.REDBRFAK, TabelaVnes.MATBR, TabelaVnes.EDINICASOS;

