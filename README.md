# PretvoriKverijaMakroaVoVBA
Претвара листа на кверија во текст фајлови во VBA *.bas фајл кој потоа може да се користи во Акцес или Ексел како таков

Се извезуваат кверијата од Акцес фајл во некоја папка со користење на Utils.ExportQuerySQL(pateka). 
Потоа се местат папките и табелите за тоа каде да ги најде кверијата и кои имиња на табели да ги замени со кои, 
наведени во тагот iminjaTabeli одвоени со „цевка“ (|) во PretvoriKverijaVoVBA.exe.config.xml

Функцијата Utils.IzveziKverijaIMakroaVoModuli(pateka) е експериментална. Ги извезува кверијата во соодветни папки со 
името на макрото во кое кверијата се повикуваат, во Акцес фајлот. Ова би значело дека во секоја од овие папки ќе треба 
да се мести нова конфигурација на претворачот, но во случај едно макро да има 30+ кверија, помага во прегледност.  

# Чекор 1 - Намести конфигурација
Во фајлот PretvoriKverijaMakroaVoVBA.exe.config.xml се местат поставки за како сакаме да се изврши извезувањето.
## VID_NA_FAJL
При извезување на кверијата од Акцес во фајлови, каков вид на фајл сакаме да биде. Основна вредност, \*.sql. Може да биде и \*.txt. Доколку се извезени како \*.sql, оваа алатка ќе ги бара тие со наставка \*.sql. 
## IME_NA_IZVEZEN_FAJL
Како сакаме да се вика фајлот на модулот што треба да се створи од самите кверија. Пример: IT1_soopstenie.bas ќе биде ВБА модул кој ќе содржи ВБА функции кои како исход враќаат „Access Jet SQL“ кверија како текст. Во суштина постоечките кверија во Акцес ги автоматизираме да бидат со динамички Ескуел код при што само имињата на табелите по потреба се менуваат, а кверито се створува во лет. 
## IME_NA_IZVEZEN_EKSEL_FAJL
Акцес наредбата DoCmd.RunSQL работи само со акциони кверија, односно кверија кои менуваат податоци (UPDATE), внесуваат податоци во табела (INSERT/SELECT INTO), бришат податоци од табела (DELETE). Кверија кои само листаат податоци нема да се извршат. За ова ќе треба да се користи класата QueryDef. Но исто така DoCmd.RunSQL работи и со извезување податоци во Ексел документ. Така SELECT може да се користи со INTO каде со посебен код се кажува во кој Ексел документ да се извезе исходот од кверито. Пример: imeXLSFajl.xls 
Во фајлот Utils.bas се наоѓа функцијата ExcelJetSQLString(imeProverka, imeXLSFajl) каде вредноста на imeProverka ќе биде име на работниот лист во новиот Ексел документ, додека imeXLSFajl ќе биде името на Ексел документот кој ќе се направи. Можат да се користат повеќе кверија едно по друго кои извезуваат во Ексел документ така што вредноста на секое квери за imeProverka мора да се разликува бидејќи во спротивно Акцес ќе врати грешка дека таков работен лист веќе постои. 
## dodajZaIzvozVoEksel
Во главно сите кверија би можеле да се прилагодат да го извезуваат исходот во Ексел документ. Ова поле служи за тоа, дали при створување на ВБА модул со динамички Ескуел код од кверијата, да се додаде делот во кверито за извезување во Ексел. Дадени кверија веќе содржат INTO збор, па треба да се прегледаат дали ќе треба во Ексел или само во локална табела во Акцес. 
По ново, само доколку кверито содржи SELECT, но не и INTO, ќе се додаде делот за извезување во Ексел, затоа што DoCmd.RunSQL извршува само акциони кверија, и кога dodajZaIzvozVoEksel е false.
## PATEKA_IZVEZEN_EKSEL_FAJL
Самите кверија кои ќе се претворат во ВБА функции, исходот при извршување може да го префрлат во Ексел документ. Ова поле е за патеката каде да се извезе тој Ексел документ со податоци од кверијата. 
## SQL_KVERIJA_PAPKA
Патека каде се наоѓаат кверијата извезени како текст фајл. 
## VBA_MAKROA_PAPKA
Патека каде се наоѓаат макроата извезени како текст фајл.
## VBA_MAKRO_VID_FAJL
Вид на фајл на модулот кој се добива како исход од претворање на кверијата во динамички Ескуел код. Обично е \*.bas, но може да биде и \*.vb.
## iminijaTabeli
За да може да се користи претворачот треба да постојат текстуални кверија. Акцес има сопствена функционалност која со
малку автоматизирање може да ги извезе сите кверија во текст фајл. За таа цел се користи функцијата Utils.ExportQuerySQL()
од ВБА модулот Utils.bas. 
Потребно е да се најдат сите табели кои се повикуваат во кверијата и соодветно да се наведе нивна замена во новиот модул кој
ќе се створи со претворачот. Тоа се прави во фајлот PretvoriKverijaMakroaVoVBA.exe.config.xml во тагот
iminijaTabeli. Пример:
```xml
	<setting name="iminijaTabeli" serializeAs="Xml">
                <value>
                    <ArrayOfString xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                        xmlns:xsd="http://www.w3.org/2001/XMLSchema">
                        <string>Q_M10|TabelaKveriM10|1</string>
                        <string>Q_M1|TabelaKveriM1|1</string>
                        <string>Q_M2|TabelaKveriM2|1</string>
                        <string>Q_M3|TabelaKveriM3|1</string>
                        <string>Q_M4|TabelaKveriM4|1</string>
                        <string>Q_M5|TabelaKveriM5|1</string>
                        <string>Q_M6|TabelaKveriM6|1</string>
                        <string>Q_M7|TabelaKveriM7|1</string>
                        <string>Q_M8|TabelaKveriM8|1</string>
                        <string>Q_M9|TabelaKveriM9|1</string>
                        <string>TabelaVnes|TabelaVnes|0</string>
                        <string>TabelaVnes2|TabelaVnes|0</string>
                        <string>TabelaAdresar|TabelaAdresar|0</string>
                        <string>TabelaAdresar1|TabelaAdresar|0</string>
                        <string>Tgodisnik_1|TabelaGodishnik|1</string>
                        <string>Tgodisnik_4|TabelaGodishnik4|1</string>
                        <string>Tgodisnik_7|TabelaGodishnik7|1</string>
                        <string>Tpol|TabelaPol|1</string>
                        <string>RZSVNESI_T_KD_SV50A_2009|TabelaVnesSV50|0</string>
                    </ArrayOfString>
                </value>
            </setting>
```
Тука TabelaKveriM10 ќе биде заменско име за табелата Q_M10, односно, во модулот ќе стане параметар на функцијата која створува дадено квери. Така името на табелата може да се менува по потреба, и секој пат ќе се створи квери со новото име на табелата, пратено како параметар на функцијата. 
Потребно е да се излистаат сите табели, и нивни заменски имиња. Заменувањето се извршува со точно поклопување, односно ако едно име се содржи во друго име, ќе се замени само ако се поклопат сите знаци од почеток до крај на зборот, додека ако се содржи во друг збор, нема да се изврши замена. За сега заменувањето на дадени табели ќе се изврши онoлку пати колку што има повторувања на име на табела во друго име на табела само за серверските табели. Ова е дефект. Ќе се додадат неколку пати истите имиња на табели како параметар на функциите во модулот. Така ќе треба рачно да се поправи и да се избришат дупликат параметрите. 
Претворачот ќе работи и доколку не се наведат имиња на табели, но тие нема да бидат заменливи.
Третиот параметар во на пример полето  <string>Q_M10|TabelaKveriM10|1</string> одвоен со цевка означува дали дадената табела е меѓу-табела или табела која постои на сервер (табела за внес, табела адресар, консулатциона датотека.) Меѓу-табелите треба да се избришат бидејќи зафаќаат простор и држат непотребни податоци. Потребно е да се избришат меѓу-табелите за да се смали големината на Акцес фајлот. Ова помага при архивирање на апликации. Бројот „1“ значи дека дадената табела е меѓу-табела и истата треба да се додаде во листата за бришење во додадената метода што се вика brishiMegjuTabeli() во новиот ВБА модул што ќе се створи. Притоа истата ќе биде додадена да се повика во главната функција на макрото. Во краен случај оваа метода, може да се повика и на копче. Добрата работа е што меѓу-табелите се излистани за да не мора ова пешки да го правиме. Така само со еден клик ќе се избришат непотребните меѓу-табели. 
# Чекор 2 - Извези кверија и макроа од Акцес во текст фајлови и ВБА модули. 
Извезување на кверија и макроа од Акцес фајлот е прилично едноставно бидејќи Акцес веќе има своја функционалност за намената.
Се прават три копчиња во Акцес форма, „Извези кверија во текст фајл“, „Извези само макроа во текст фајл“ и „Извези кверија и макроа во текст фајл“. 
Притоа на клик, треба да се повикаат фунцкиите Utils.ExportQuerySQL(), Utils.PretvoriMakroaVoModuliISnimiNaDisk() и Utils.IzveziKverijaIMakroaVoModuli(), соодветно, како на сликата:
![Извези кверија и макроа од Акцес во текст фајлови и ВБА модули](https://github.com/bluePlayer/PretvoriKverijaMakroaVoVBA/blob/master/PretvoriKverijaMakroaVoVBA/sliki/PretvoriKverijaMakroaVoVBA1.png)

Во ВБА тоа би изгледало од прилика вака:

```vb
Private Sub izveziKverijaVoTekstBtn_Click()
    Utils.ExportQuerySQL (Konstanti.KVERIJA_PATEKA_ZA_IZVEZUVANJE)
End Sub

Private Sub izveziMakroaIKverijaVoTekstBtn_Click()
    Utils.IzveziKverijaIMakroaVoModuli (Konstanti.KVERIJA_I_MAKROA_PATEKA_ZA_IZVEZUVANJE)
End Sub

Private Sub izveziMakroaVoTekstBtn_Click()
    Utils.PretvoriMakroaVoModuliISnimiNaDisk (Konstanti.MAKROA_PATEKA_ZA_IZVEZUVANJE)
End Sub
```

Каде Konstanti.bas е модул во Акцес кој ги содржи константите KVERIJA_PATEKA_ZA_IZVEZUVANJE, KVERIJA_I_MAKROA_PATEKA_ZA_IZVEZUVANJE и MAKROA_PATEKA_ZA_IZVEZUVANJE. од прилика вака:

```vb
Public Const KVERIJA_PATEKA_ZA_IZVEZUVANJE As String = "C:\Users\user1\Documents\output\kverija\"
Public Const MAKROA_PATEKA_ZA_IZVEZUVANJE As String = "C:\Users\user1\Documents\output\makroa\"
Public Const KVERIJA_I_MAKROA_PATEKA_ZA_IZVEZUVANJE As String = "C:\Users\user1\Documents\output\kverija_makroa\"
```
- Функцијата Utils.ExportQuerySQL (Konstanti.KVERIJA_PATEKA_ZA_IZVEZUVANJE) ги извезува само кверијата од Акцес во текст фајл со соодветното име. 
- Функцијата Utils.IzveziKverijaIMakroaVoModuli (Konstanti.KVERIJA_I_MAKROA_PATEKA_ZA_IZVEZUVANJE) ги извезува макроата во ВБА модули како засебен фајл на диск со соодветното име, но претходно, секое макро го претвора во ВБА модул во самиот Акцес со претставка „Converted macro- “. При секое претворање се појавува мал дијалог прозор на кој се притиска Ентер и не може да се избегне. Но, ова се прави само еднаш, и откако ќе се извезат сите макроа нема потреба да се прави истото одпочеток. 
- Функцијата Utils.PretvoriMakroaVoModuliISnimiNaDisk (Konstanti.MAKROA_PATEKA_ZA_IZVEZUVANJE) ги ствара ВБА модулите со име „Converted macro- imeNaMakro“ во Акцес (се појавува дијалог прозор на кој се притиска Ентер.), додава папка на диск со името на макрото, го снима ВБА модулот на макрото „imeNaMakro.bas“ во истоимената папка, и потоа со овој модул се служи да најде кои кверија му припаѓаат на макрото, така што модулот кој ќе се створи од кверијата ќе се сними во соодветната папка на макрото. Ова многу го олеснува процесот да се претвори макро во модул со динамички ескуел код. За сега експериментален код. 
- Постои продобрување на кодот, така што во ВБА модулот на диск наредбата
```vb
DoCmd.OpenQuery imeNaKveri, acNormal, acEdit 
```
ќе се замени со 
```vb
DoCmd.RunSQL imeNaKveri_sql(lista na tabeli)
```
Притоа, главниот код на извезеното макро ќе се вметне во новиот модул, со заменети „OpenQuery imeNaKveri“  наредби со соодветни „DoCmd.RunSQL imeNaKveri“. Во случај макрото да има 40 кверија, ова ќе го олесни автоматизирањето на макрото. 
Соодветно, во новиот модул ќе се додаде и функцијата brishiMegjuTabeli() која ќе ги содржи меѓу-табелите на даденото макро кои треба да се избришат после нивна употреба. 

# Чекор 3 - Користење
- Претворачот е конзолна апликација. Доколки сите ставки се наместени како што треба, се извршува само со двоен клик или во командна линија, со влез во папката каде се наоѓа PretvoriKverijaMakroaVoVBA.exe. При извршување, прво ќе се створи модул за самите кверија. Колку што има кверија во наведената папка толку функции ќе се додадат во модулот. При извршување претворачот прави некое основно форматирање, па ако се многу кверија, кои се прилично долги, крајниот фајл може да биде неколку илјади линии. Затоа е добро да се поделат кверијата во една папка кои се повикуваат во дадено макро. 
Но, ова може да биде долга и здодевна задача, особено ако се многу макроа со многу кверија во себе. 
За таа цел, вториот дел на извршување, откако ќе се притисне копче на тастатурата, се наоѓаат макроата по име и во нив се бараат керијата по име. Ако макроата извезени од претходно со претставка („Converted macro- “) содржат име на дадено на квери во себе, тоа ќе се сними во папка со име на макрото. Притоа ќе се створи нов модул со исто име на макрото каде ќе бидат кверијата како динамички ескуел код. Овој дел е експериментален и потребно е веќе во папката на макрото, да постои модулот на макрото извезен од претходно како \*.bas фајл, пример Converted macro- imeNaMakro.bas. Но, извезувањето на макроата од Акцес во ВБА модул се прави автоматски, со повикување на функцијата Utils.PretvoriMakroaVoModuliISnimiNaDisk(). 
Претворањето од прилика изгледа вака:
![Користење на претворачот](https://github.com/bluePlayer/PretvoriKverijaMakroaVoVBA/blob/master/PretvoriKverijaMakroaVoVBA/sliki/PretvoriKverijaMakroaVoVBA2.png)

Во случајов, вториот дел, претворање на макроа и кверија во ВБА модули со динамички ескуел код не успеа затоа што во папката на макрото Macro5 не постоеше ВБА модул со име, Macro5.bas. Со извезување на истиот, ќе се створи нов модул со име Macro5_module.bas кој ќе ги содржи кверијата како динамички ескуел, притоа, е модул кој ги содржи само кверијата кои се повикуваат во истоименото макро во Акцес, и се наоѓа во своја соодветна папка.
![Користење на претворачот - 2](https://github.com/bluePlayer/PretvoriKverijaMakroaVoVBA/blob/master/PretvoriKverijaMakroaVoVBA/sliki/PretvoriKverijaMakroaVoVBA3.png)

- Macro5.bas изгледа вака

```vb
Option Compare Database

'------------------------------------------------------------
' Macro5
'
'------------------------------------------------------------
Function Macro5()
On Error GoTo Macro5_Err

    DoCmd.SetWarnings False
    DoCmd.OpenQuery "Regioni_0", acViewNormal, acEdit
    DoCmd.OpenQuery "Regioni_01", acViewNormal, acEdit
    DoCmd.OpenReport "Regioni", acViewPreview, "", ""


Macro5_Exit:
    Exit Function

Macro5_Err:
    MsgBox Error$
    Resume Macro5_Exit

End Function
```
додека модулот кој ќе се створи во папката Macro5, Macro5_module.bas изгледа вака:
```vb
Public Function Regioni_0_sql(ByVal TabelaVnes As String , ByVal TabelaNasMesta As String ) As String
  Dim sql as String
  sql = sql & "SELECT ""[V-50 "" AS OBRAZEC, " & vbNewLine
  sql = sql & "" & TabelaVnes & ".GOD, " & vbNewLine
  sql = sql & """0"" AS REGIONID, " & vbNewLine
  sql = sql & """Republika Makedonija"" AS REGION, " & vbNewLine
  sql = sql & "Count(" & TabelaVnes & ".OBRAZEC) AS SE, " & vbNewLine
  sql = sql & "Sum(IIf([pol]=""1"",1,0)) AS Mazi, " & vbNewLine
  sql = sql & "Sum(IIf([pol]=""2"",1,0)) AS Zeni, " & vbNewLine
  sql = sql & "" & TabelaVnes & ".ZDRZAVA INTO Regioni_1 " & vbNewLine
  sql = sql & "FROM " & TabelaVnes & " LEFT JOIN " & TabelaNasMesta & " ON " & TabelaVnes & ".ZNASMES = " & TabelaNasMesta & ".NASID" & vbNewLine
  sql = sql & "GROUP BY ""[V-50 "", " & TabelaVnes & ".GOD, ""0"", ""Republika Makedonija"", " & TabelaVnes & ".ZDRZAVA" & vbNewLine
  sql = sql & "HAVING (((" & TabelaVnes & ".ZDRZAVA)=""807""));" & vbNewLine
  sql = sql & "" & vbNewLine
  sql = sql & "" & vbNewLine
  Regioni_0_sql = sql
End Function

Public Function Regioni_01_sql(ByVal TabelaVnes As String , ByVal TabelaNasMesta As String ) As String
  Dim sql as String
  sql = sql & "INSERT INTO Regioni_1 ( OBRAZEC, GOD, REGIONID, REGION, SE, Mazi, Zeni, ZDRZAVA )" & vbNewLine
  sql = sql & "SELECT ""[V-50 "" AS OBRAZEC, " & vbNewLine
  sql = sql & "" & TabelaVnes & ".GOD, " & vbNewLine
  sql = sql & "" & TabelaNasMesta & ".REGIONID, " & vbNewLine
  sql = sql & "" & TabelaNasMesta & ".REGION, " & vbNewLine
  sql = sql & "Count(" & TabelaVnes & ".OBRAZEC) AS CountOfOBRAZEC, " & vbNewLine
  sql = sql & "Sum(IIf([pol]=""1"",1,0)) AS Mazi, " & vbNewLine
  sql = sql & "Sum(IIf([pol]=""2"",1,0)) AS Zeni, " & vbNewLine
  sql = sql & "" & TabelaVnes & ".ZDRZAVA " & vbNewLine
  sql = sql & "FROM " & TabelaVnes & " INNER JOIN " & TabelaNasMesta & " ON " & TabelaVnes & ".ZNASMES = " & TabelaNasMesta & ".NASID" & vbNewLine
  sql = sql & "GROUP BY ""[V-50 "", " & TabelaVnes & ".GOD, " & TabelaNasMesta & ".REGIONID, " & TabelaNasMesta & ".REGION, " & TabelaVnes & ".ZDRZAVA" & vbNewLine
  sql = sql & "HAVING (((" & TabelaVnes & ".ZDRZAVA)=""807""));" & vbNewLine
  sql = sql & "" & vbNewLine
  sql = sql & "" & vbNewLine
  Regioni_01_sql = sql
End Function

'------------------------------------------------------------
' Macro5
'
'------------------------------------------------------------
Function Macro5()
On Error GoTo Macro5_Err

    DoCmd.SetWarnings False
    DoCmd.RunSQL Macro5_module.Regioni_0(TabelaVnes , TabelaNasMesta)
    DoCmd.RunSQL Macro5_module.Regioni_01(TabelaVnes , TabelaNasMesta)
    DoCmd.RunSQL Macro5_module.Regioni(TabelaVnes , TabelaNasMesta)

	'Call brishiMegjuTabeli
	
Macro5_Exit:
    Exit Function

Macro5_Err:
    MsgBox Error$
    Resume Macro5_Exit

End Function

Public Sub brishiMegjuTabeli()
DoCmd.RunSQL "drop table TabelaNasMesta"
End Sub

```
каде се вметнаа главната функција на макрото, со повикување на кверијата, како и методата за бришење на меѓу-табелите, каде во случајов само TabelaNasMesta е излистана бидејќи е меѓу-табела, додека TabelaVnes е серверска табела.
# Чекор 4 - Исправки на створените фајлови и нивна употреба 
- Бидејќи претворачот може неколку пати да најде имиња на табели кои се содржат во други имиња на табели, ова значи дека во потписот на ВБА методите ќе се додадат дупликат параметри кои треба да се избришат пешки бидејќи засега, овој проблем не е решен. Функција во ВБА не може да има два параметри со исто име. 
- Се случува дадено квери да е доста долго, па иако има код за да ја расцепи линијата на две или повеќе пократки линии, сепак може да прекорачи одреден број на букви во линија. Притоа, оваа линија мора да се најде и расцепи рачно, за да го искомпајлира Акцес фајлот новиот код. Треба ескуел кодот кој се добива како исход од функцијат да биде важечки ескуел код за да може да го изврши DoCmd.RunSQL наредбата. 
- Новиот ВБА модул се додава во Акцес со наредбата Database tools -> Visual Basic Editor -> Modules -> Import file -> ime_na_vba_modul_module.bas. Притоа мора задожително да се искомпајлира овој нов фајл со наредбата Debug -> Compile. Доколку не се искомпајлира, и се појави порака за грешка, мора да се поправат сите грешки се додека не се искомпајлира новиот фајл. Повеќе макроа значат повеќе ВБА модули. Истите треба да се додадат во Акцес фајлот, поправат грешките во нив и да се искомпајлира Акцес фајлот со овие нови фајлови. Во примерот погоре, на клик-евент на копче би се повикала функцијата Macro5(). Притоа ако сме сигурни дека се е во ред, ќе ја одкоментираме линијата со „Call brishiMegjuTabeli“ и ќе искомпајлираме повторно. Така секој пат, ВБА модулот ќе ги брише меѓу-табелите откако ќе се добие тоа што сакаме, се разбира доколку не се брише табелата во која се наоѓа крајниот исход на макрото, пример излезна табела за објавување. 
