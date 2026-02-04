using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PretvoriKverijaMakroaVoVBA
{
    public class AppState
    {
        public string vidNaFajl { get; set; }
        public string imeNaIzvezenFajl { get; set; }
        public string imeNaIzvezenEkselFajl { get; set; }
        public string patekaNaIzvezenEkselFajl { get; set; }
        public string sqlKverijaPapka { get; set; }
        public string vbaMakroaPapka { get; set; }
        public string vbaMakroVidFajl { get; set; }
        public string imeFajlTabeliKonstanti { get; set; }

        public List<ImeTabelaZamena> iminjaTabeliZameni { get; set; }

        public List<VbaMakro> listaVbaMakroa { get; set; }

        public AppState()
        {
            vidNaFajl = Properties.Settings.Default.VID_NA_FAJL;
            imeNaIzvezenFajl = Properties.Settings.Default.IME_NA_IZVEZEN_FAJL;
            imeNaIzvezenEkselFajl = Properties.Settings.Default.IME_NA_IZVEZEN_EKSEL_FAJL;
            patekaNaIzvezenEkselFajl = Properties.Settings.Default.PATEKA_IZVEZEN_EKSEL_FAJl;
            sqlKverijaPapka = Properties.Settings.Default.SQL_KVERIJA_PAPKA;
            vbaMakroaPapka = Properties.Settings.Default.VBA_MAKROA_PAPKA;
            vbaMakroVidFajl = Properties.Settings.Default.VBA_MAKRO_VID_FAJL;
            imeFajlTabeliKonstanti = Properties.Settings.Default.IME_FAJL_TABELI_KONSTANTI;

            iminjaTabeliZameni = new List<ImeTabelaZamena>();

            VchitajIminjaTabeliZameni();

            listaVbaMakroa = new List<VbaMakro>();
        }

        public void VchitajIminjaTabeliZameni()
        {
            List<string> tabeli = new List<string>();

            foreach (string tbl in Properties.Settings.Default.iminijaTabeli)
            {
                string[] parTabeli = tbl.Split(Konstanti.ODDELUVACH_ZA_IMINJA_NA_TABELI);

                if (!tabeli.Contains(parTabeli[1]))
                {
                    ImeTabelaZamena imaTabelaZamena = new ImeTabelaZamena();
                    imaTabelaZamena.ime = parTabeli[0];
                    imaTabelaZamena.imeZamena = parTabeli[1];
                    imaTabelaZamena.daliEMegjuTabela = parTabeli[2] == "1";

                    iminjaTabeliZameni.Add(imaTabelaZamena);
                    tabeli.Add(parTabeli[1]);
                }
            }
        }

        public void IzveziIminjaTabeliKonstantiVoFajl()
        {
            List<string> zamenaIminjaTabeli = new List<string>();

            foreach (ImeTabelaZamena tbl in iminjaTabeliZameni)
                zamenaIminjaTabeli.Add(tbl.imeZamena);

            foreach (ImeTabelaZamena tbl in iminjaTabeliZameni)
                if (!tbl.daliEMegjuTabela)
                {
                    zamenaIminjaTabeli.Add(tbl.imeZamena + "_TEST");
                }

            // gi stava konstantite pred da se vmetne kodot na glavnata makro-funkcija
            StringBuilder iminjaTabeliKonstanti = new StringBuilder();

            foreach (string imeTabelaZamena in zamenaIminjaTabeli)
            {
                iminjaTabeliKonstanti.Append("Public Const " + Konstanti.PRETSTAVKA_IME_TABELA_KONSTANTA + imeTabelaZamena.ToUpper() + " As String = \"" + imeTabelaZamena + "\"" + Environment.NewLine);
            }

            iminjaTabeliKonstanti.Append(Environment.NewLine);

            File.WriteAllText(sqlKverijaPapka + "\\..\\" + imeFajlTabeliKonstanti, iminjaTabeliKonstanti.ToString());
        }

        /// <summary>
        /// Pretvora lista na kverija koi se vo tekstualen format vo dadena papka, vo VBA modul.
        /// Ako ima 100 kverija, kje gi pretvori vo eden VBA modul so 100 funkcii koi vrakjaat SQL string 
        /// (dinamichki SQL). 
        /// </summary>
        public void IzvrshiPretvoranje()
        {
            try
            {
                // Primeri pr = new Primeri();

                DirectoryInfo sqlKverijaDir = new DirectoryInfo(sqlKverijaPapka);
                FileInfo[] sqlKverijaFajlovi = sqlKverijaDir.GetFiles(vidNaFajl);
                StringBuilder newFileBuilder = new StringBuilder();
                StringBuilder povikKverijaFunkcijaSB = new StringBuilder();
                string fileName = imeNaIzvezenFajl;

                DirectoryInfo vbaMakroaKverijaDir = new DirectoryInfo(vbaMakroaPapka);
                FileInfo[] vbaMakroaKverijaFajlovi = vbaMakroaKverijaDir.GetFiles(vbaMakroVidFajl);

                // dodaj i otvori nova procedura koja gi povikuva kverijata vnatre
                int brojKveri = 1;
                povikKverijaFunkcijaSB.Append("Public Sub IzvrshiKverija()" + Environment.NewLine);
                povikKverijaFunkcijaSB.Append("    Dim sql as String" + Environment.NewLine);
                povikKverijaFunkcijaSB.Append("    DoCmd.SetWarnings False" + Environment.NewLine);

                foreach (FileInfo file in sqlKverijaFajlovi)
                {
                    string contents = File.ReadAllText(file.FullName);

                    string sqlString =
                        PretvoriJetSQLKverijaVoVBAUtils.Pretvori(
                            file.Name.Replace(".", "_"),
                            contents,
                            iminjaTabeliZameni,
                            null,
                            null,
                            null);

                    newFileBuilder.Append(sqlString);

                    string kveriMetoda =
                        sqlString.Split('\n')[0]
                            .Replace("Public Function", string.Empty)
                            .Replace("ByVal", string.Empty)
                            .Replace("As String", string.Empty);

                    povikKverijaFunkcijaSB.Append("    '-----" + brojKveri.ToString() + "-----" + Environment.NewLine);
                    povikKverijaFunkcijaSB.Append("    sql = " + kveriMetoda + Environment.NewLine);
                    povikKverijaFunkcijaSB.Append("    Debug.Print sql" + Environment.NewLine);
                    povikKverijaFunkcijaSB.Append("    DoCmd.RunSQL (sql)" + Environment.NewLine);
                    povikKverijaFunkcijaSB.Append(Environment.NewLine);

                    brojKveri += 1;

                    Console.WriteLine("Pretvoriv: " + file.Name + " vo VBA kod! ");
                }

                // zatvori procedura koja gi povikuva kverijata vnatre 
                povikKverijaFunkcijaSB.Append("    DoCmd.SetWarnings True" + Environment.NewLine);
                povikKverijaFunkcijaSB.Append("End Sub" + Environment.NewLine);

                foreach (ImeTabelaZamena tbl in iminjaTabeliZameni)
                {
                    povikKverijaFunkcijaSB = povikKverijaFunkcijaSB.Replace(
                        tbl.imeZamena,
                        Properties.Settings.Default.IME_FAJL_TABELI_KONSTANTI.Replace("bas", string.Empty) + tbl.imeZamena);
                }

                File.WriteAllText(Properties.Settings.Default.SQL_KVERIJA_PAPKA + "\\" + fileName, newFileBuilder.ToString());
                File.AppendAllText(Properties.Settings.Default.SQL_KVERIJA_PAPKA + "\\" + fileName, povikKverijaFunkcijaSB.ToString());

                Console.WriteLine("Zavrshiv so pretvaranje na kverijata vo VBA kod! ");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Greshka: " + ex.Message);
            }
        }

        /// <summary>
        /// Pretvora kverija vo VBA modul zavisno od postoechki papki na makroa. 
        /// Ako dadeno kveri ima ime koe se sodrzhi vo lista na kverija koi se povikuvaat
        /// od dadeno makro, istoto kveri kje se pretvori vo funkcija vo VBA modul (modulot ima ime kako
        /// makroto). Ova znachi deka vo papkata so ime na makroto, kje postoi VBA modul vo koj ima
        /// kverija/methodi so iminja na kverijata koi se povikuvaat vo samoto makro. 
        /// No potrebno e da postojat makroata izvezeni od prethodno kako VBA modul od samiot Akces fajl. 
        /// Taka noviot modul kje se vika ime_na_makro_SQL.bas, so nastavkata "_SQL". 
        /// </summary>
        public void IzvrshiPretvoranjeMakroaIKverija()
        {
            try
            {
                DirectoryInfo sqlMakroaDir = new DirectoryInfo(vbaMakroaPapka);
                DirectoryInfo[] sqlKverijaDir = sqlMakroaDir.GetDirectories();

                foreach (DirectoryInfo di in sqlKverijaDir)
                {
                    FileInfo[] sqlKverijaFajlovi = di.GetFiles(vidNaFajl);
                    StringBuilder newFileBuilder = new StringBuilder();

                    string fileName = di.Name;
                    string moduleFileName = string.Empty;

                    Console.WriteLine("Rabotam vo papka: " + di.Name);

                    moduleFileName = fileName + "_SQL.bas";

                    newFileBuilder.Append("Option Compare Database" + Environment.NewLine + Environment.NewLine);

                    // gi pretvora kverijata vo funkcii
                    foreach (FileInfo file in sqlKverijaFajlovi)
                    {
                        string contents = File.ReadAllText(file.FullName);

                        string sqlString =
                            PretvoriJetSQLKverijaVoVBAUtils.Pretvori(
                                file.Name.Replace(".", "_"),
                                contents,
                                iminjaTabeliZameni,
                                null,
                                null,
                                null);

                        newFileBuilder.Append(sqlString);
                        Console.WriteLine("Pretvoriv: " + file.Name + " vo VBA kod! ");
                    }

                    File.WriteAllText(di.FullName + "\\" + moduleFileName, newFileBuilder.ToString());

                    FileInfo[] vbaModuliFajlovi = di.GetFiles("*.bas").OrderBy(fi => fi.Name).ToArray();

                    string makroSodrzhina = File.ReadAllText(vbaModuliFajlovi[0].FullName);
                    string[] makroLinii = makroSodrzhina.Split('\n');

                    string modulSodrzhina = File.ReadAllText(vbaModuliFajlovi[1].FullName);
                    string[] modulLinii = modulSodrzhina.Split('\n');

                    // ostanati linii na makroto
                    StringBuilder writeLines = new StringBuilder();

                    // TODO da se dovrshi ova, pechati po povekje pati edno isto
                    foreach (string makroLinija in makroLinii)
                    {
                        if (makroLinija.Contains("SetWarnings"))
                        {
                            writeLines.Append("    Dim sql as String" + Environment.NewLine);
                        }

                        if (!makroLinija.Contains("OpenQuery"))
                        {
                            if (!makroLinija.Contains("Option Compare Database"))
                                writeLines.Append(makroLinija + Environment.NewLine);
                        }
                        else
                        {
                            int brojRed = 1;

                            foreach (FileInfo file in sqlKverijaFajlovi)
                            {
                                foreach (string modulLinija in modulLinii)
                                {
                                    if (modulLinija.Contains("Public Function " + file.Name.Replace(".", "_")))
                                    {
                                        string metodPotpis =
                                            modulLinija
                                                .Replace("Public Function ", "")
                                                .Replace("ByVal", "")
                                                .Replace("As String", "");

                                        // gi zamenuva parametrite vo povikot na funkciite so iminjata na konstantite
                                        int pochetok = metodPotpis.IndexOf('(');
                                        int kraj = metodPotpis.IndexOf(')');
                                        string del1 = metodPotpis.Substring(0, pochetok);
                                        string del2 = metodPotpis.Substring(pochetok);

                                        foreach (ImeTabelaZamena tbl in iminjaTabeliZameni)
                                        {
                                            del2 = del2.Replace(tbl.imeZamena, imeFajlTabeliKonstanti.Replace("bas", string.Empty) + Konstanti.PRETSTAVKA_IME_TABELA_KONSTANTA + tbl.imeZamena.ToUpper());
                                        }

                                        metodPotpis = del1 + del2;

                                        string modulIMetoda = fileName + "_SQL." + metodPotpis;

                                        if (!writeLines.ToString().Contains(modulIMetoda))
                                        {
                                            writeLines.Append("' ----- " + brojRed.ToString() + " ----- " + Environment.NewLine);
                                            writeLines.Append("    sql = " + modulIMetoda + Environment.NewLine);
                                            writeLines.Append("    Debug.Print sql" + Environment.NewLine);
                                            writeLines.Append("    DoCmd.RunSQL sql" + Environment.NewLine);
                                            writeLines.Append(Environment.NewLine);

                                            brojRed += 1;
                                        }
                                    }
                                }
                            }

                            if (!writeLines.ToString().Contains("Call brishiMegjuTabeli()"))
                            {
                                writeLines.Append("    'Call brishiMegjuTabeli()" + Environment.NewLine + Environment.NewLine);
                            }
                        }
                    }

                    File.AppendAllText(di.FullName + "\\" + moduleFileName, writeLines.ToString() + Environment.NewLine);

                    StringBuilder dropTablesSB = new StringBuilder();

                    dropTablesSB.Append("Public Sub brishiMegjuTabeli()" + Environment.NewLine);

                    for (int i = 0; i < iminjaTabeliZameni.Count; i += 1)
                    {
                        if (iminjaTabeliZameni[i].daliEMegjuTabela && newFileBuilder.ToString().Contains(iminjaTabeliZameni[i].imeZamena))
                            dropTablesSB.Append("    DoCmd.RunSQL \"drop table \" & " + Properties.Settings.Default.IME_FAJL_TABELI_KONSTANTI.Replace("bas", string.Empty) + Konstanti.PRETSTAVKA_IME_TABELA_KONSTANTA + iminjaTabeliZameni[i].imeZamena.ToUpper() + Environment.NewLine);
                    }

                    dropTablesSB.Append("End Sub" + Environment.NewLine);

                    File.AppendAllText(di.FullName + "\\" + moduleFileName, dropTablesSB.ToString());

                    Console.WriteLine("Zavrshiv so pretvaranje na kverijata vo VBA kod, fajl: " + moduleFileName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("IzvrshiPretvoranjeMakroaIKverija(): " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : string.Empty));
            }

            Console.ReadLine();
        }

        // TODO da se zamenat zakucanite iminja na tabeli so tie od konfiguracija
        public void DodajTestTabeli()
        {
            List<string> pomIminjaTabeli = new List<string>();
            string imeTabela = string.Empty;
            string imeTestTabela = string.Empty;
            StringBuilder ishod = new StringBuilder();

            ishod.Append(Environment.NewLine);

            ishod.Append("Public Sub DodajTestTabeli()" + Environment.NewLine);
            ishod.Append("    Dim sql as String" + Environment.NewLine);
            ishod.Append("    DoCmd.SetWarnings False" + Environment.NewLine);
            ishod.Append(Environment.NewLine);

            foreach (ImeTabelaZamena tbl in iminjaTabeliZameni)
            {
                if (!pomIminjaTabeli.Contains(tbl.imeZamena))
                {
                    pomIminjaTabeli.Add(tbl.imeZamena);

                    if (!tbl.daliEMegjuTabela)
                    {
                        imeTabela = imeFajlTabeliKonstanti.Split('.')[0] + "." + Konstanti.PRETSTAVKA_IME_TABELA_KONSTANTA + tbl.imeZamena.ToUpper();
                        imeTestTabela = imeTabela + "_TEST";
                        ishod.Append("    sql = \"select \" & " + imeTabela + " & \".* into \" & " + imeTestTabela + " & \" from \" & " + imeTabela + " " + Environment.NewLine);
                        ishod.Append("    Debug.Print sql" + Environment.NewLine);
                        ishod.Append("    DoCmd.RunSQL (sql) " + Environment.NewLine);
                        ishod.Append(Environment.NewLine);
                    }
                }
            }

            ishod.Append("    DoCmd.SetWarnings True" + Environment.NewLine);
            ishod.Append("End Sub" + Environment.NewLine);

            // ----- dodaj kod za brishenje test tabeli -----
            pomIminjaTabeli = new List<string>();

            ishod.Append(Environment.NewLine);
            ishod.Append("Public Sub brishiMegjuTabeli()" + Environment.NewLine);
            ishod.Append("    Dim sql as String" + Environment.NewLine);
            ishod.Append("    DoCmd.SetWarnings False" + Environment.NewLine);
            ishod.Append(Environment.NewLine);

            foreach (ImeTabelaZamena tbl in iminjaTabeliZameni)
            {
                if (!pomIminjaTabeli.Contains(tbl.imeZamena))
                {
                    pomIminjaTabeli.Add(tbl.imeZamena);

                    if (!tbl.daliEMegjuTabela)
                    {
                        imeTabela = imeFajlTabeliKonstanti.Split('.')[0] + "." + Konstanti.PRETSTAVKA_IME_TABELA_KONSTANTA + tbl.imeZamena.ToUpper();
                        imeTestTabela = imeTabela + "_TEST";
                        ishod.Append("    If Utils.daliTabelataPostoi(" + imeTestTabela + ", CurrentDb) Then " + Environment.NewLine);
                        ishod.Append("        sql = \"drop table \" & " + imeTestTabela + " " + Environment.NewLine);
                        ishod.Append("        Debug.Print sql" + Environment.NewLine);
                        ishod.Append("        DoCmd.RunSQL sql" + Environment.NewLine);
                        ishod.Append("    End If");
                        ishod.Append(Environment.NewLine);
                        ishod.Append(Environment.NewLine);
                    }
                }
            }

            ishod.Append("    DoCmd.SetWarnings True" + Environment.NewLine);
            ishod.Append("End Sub" + Environment.NewLine);

            File.AppendAllText(sqlKverijaPapka + "\\" + imeNaIzvezenFajl, ishod.ToString());

            Console.WriteLine("Zavrshiv so VBA kod za test tabeli! ");
            Console.ReadLine();
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("vidNaFajl: " + vidNaFajl + Environment.NewLine);
            sb.Append("imeNaIzvezenFajl: " + imeNaIzvezenFajl + Environment.NewLine);
            sb.Append("imeNaIzvezenEkselFajl: " + imeNaIzvezenEkselFajl + Environment.NewLine);
            sb.Append("patekaNaIzvezenEkselFajl: " + patekaNaIzvezenEkselFajl + Environment.NewLine);
            sb.Append("sqlKverijaPapka: " + sqlKverijaPapka + Environment.NewLine);
            sb.Append("vbaMakroaPapka: " + vbaMakroaPapka + Environment.NewLine);
            sb.Append("vbaMakroVidFajl: " + vbaMakroVidFajl + Environment.NewLine);
            sb.Append("imeFajlTabeliKonstanti: " + imeFajlTabeliKonstanti + Environment.NewLine);

            foreach (ImeTabelaZamena imeTblZamena in iminjaTabeliZameni)
                sb.Append(imeTblZamena.ToString());

            foreach (VbaMakro vbaMakro in listaVbaMakroa)
                sb.Append(vbaMakro.ToString());

            return sb.ToString();
        }
    }
}
