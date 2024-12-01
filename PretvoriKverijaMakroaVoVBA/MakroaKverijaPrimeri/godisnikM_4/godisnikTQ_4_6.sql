INSERT INTO Tgodisnik_4 ( GOD, sort, NAZIV, POL, pol_naziv, KOL1, KOL1A, KOL2, KOL3, KOL4, KOL5, KOL6, KOL7, KOL8, KOL9 )
SELECT TabelaVnes.GOD, [GODISNIK] & [ID1_GOD] AS sort, TabelaAdresar.VIDSKOLA AS NAZIV, "" AS POL, "SE" AS pol_naziv, Count(TabelaVnes.POL) AS KOL1, Sum(IIf([DRZAVJ]="807",1,0)) AS KOL1A, Sum(IIf([DRZAVJ]="807" And [NACPRIP]="01",1,0)) AS KOL2, Sum(IIf([DRZAVJ]="807" And [NACPRIP]="03",1,0)) AS KOL3, Sum(IIf([DRZAVJ]="807" And [NACPRIP]="06",1,0)) AS KOL4, Sum(IIf([DRZAVJ]="807" And [NACPRIP]="05",1,0)) AS KOL5, Sum(IIf([DRZAVJ]="807" And [NACPRIP]="04",1,0)) AS KOL6, Sum(IIf([DRZAVJ]="807" And [NACPRIP]="27",1,0)) AS KOL7, Sum(IIf([DRZAVJ]="807" And ([NACPRIP]<>"01" And [NACPRIP]<>"03" And [NACPRIP]<>"06" And [NACPRIP]<>"04" And [NACPRIP]<>"05" And [NACPRIP]<>"27"),1,0)) AS KOL8, Sum(IIf([DRZAVJ]<>"807",1,0)) AS KOL9
FROM TabelaVnes INNER JOIN TabelaAdresar ON (TabelaVnes.OTSEK = TabelaAdresar.SIFRAOTSEK) AND (TabelaVnes.EVS = TabelaAdresar.EDINICASOS) AND (TabelaVnes.MATBR = TabelaAdresar.MATBR)
WHERE (((TabelaAdresar.GODISNIK)<>"3"))
GROUP BY TabelaVnes.GOD, [GODISNIK] & [ID1_GOD], TabelaAdresar.VIDSKOLA, "", "SE";

