using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace PretvoriKverijaMakroaVoVBA
{
    public class PretvoriJetSQLKverijaVoVBAUtils
    {
        public const char SPEC_KARAKTER_ZA_ZAMENA = '@';
        public const char ODDELUVACH_ZA_IMINJA_NA_TABELI = '|';
        public const string PRETSTAVKA_IME_TABELA_KONSTANTA = "TBL_";
        public const string REGEX_TOCHNO_POKLOPUVANJE = @"\b{0}\b";

        /// <summary>
        /// Pretvora kverija vo VBA modul zavisno od postoechki papki na makroa. 
        /// Ako dadeno kveri ima ime koe se sodrzhi vo lista na kverija koi se povikuvaat
        /// od dadeno makro, istoto kveri kje se pretvori vo funkcija vo VBA modul (modulot ima ime kako
        /// makroto). Ova znachi deka vo papkata so ime na makroto, kje postoi VBA modul vo koj ima
        /// kverija/methodi so iminja na kverijata koi se povikuvaat vo samoto makro. 
        /// No potrebno e da postojat makroata izvezeni od prethodno kako VBA modul od samiot Akces fajl. 
        /// Taka noviot modul kje se vika ime_na_makro_module.bas, so nastavkata "_module". 
        /// </summary>
        public static void IzvrshiPretvoranjeMakroaIKverija()
        {
            try
            {
                DirectoryInfo sqlMakroaDir = new DirectoryInfo(Properties.Settings.Default.VBA_MAKROA_PAPKA);
                DirectoryInfo[] sqlKverijaDir = sqlMakroaDir.GetDirectories();

                foreach (DirectoryInfo di in sqlKverijaDir)
                {
                    FileInfo[] sqlKverijaFajlovi = di.GetFiles(Properties.Settings.Default.VID_NA_FAJL);
                    StringBuilder newFileBuilder = new StringBuilder();

                    string fileName = di.Name;
                    string moduleFileName = string.Empty;

                    List<string> iminijaTabeli = new List<string>();
                    List<string> zamenaIminjaTabeli = new List<string>();
                    List<bool> daliEMegjuTabelaLista = new List<bool>();

                    foreach (string tbl in Properties.Settings.Default.iminijaTabeli)
                    {
                        string[] parTabeli = tbl.Split(PretvoriJetSQLKverijaVoVBAUtils.ODDELUVACH_ZA_IMINJA_NA_TABELI);
                        iminijaTabeli.Add(parTabeli[0]);
                        zamenaIminjaTabeli.Add(parTabeli[1]);
                        daliEMegjuTabelaLista.Add(parTabeli[2] == "1");
                    }

                    Console.WriteLine("Rabotam vo papka: " + di.Name);

                    moduleFileName = fileName + "_module.bas";

                    // gi stava konstantite pred da se vmetne kodot na glavnata makro-funkcija
                    StringBuilder iminjaTabeliKonstanti = new StringBuilder();

                    foreach (string imeTabelaZamena in zamenaIminjaTabeli)
                    {
                        iminjaTabeliKonstanti.Append("Public Const " + PRETSTAVKA_IME_TABELA_KONSTANTA + imeTabelaZamena.ToUpper() + " As String = \"" + imeTabelaZamena + "\"" + Environment.NewLine);
                    }

                    iminjaTabeliKonstanti.Append(Environment.NewLine);

                    File.WriteAllText(di.FullName + "\\" + moduleFileName, iminjaTabeliKonstanti.ToString());

                    // gi pretvora kverijata vo funkcii
                    foreach (FileInfo file in sqlKverijaFajlovi)
                    {
                        string contents = File.ReadAllText(file.FullName);

                        string sqlString =
                            PretvoriJetSQLKverijaVoVBAUtils.Pretvori(
                                file.Name.Replace(".", "_"),
                                contents,
                                iminijaTabeli,
                                zamenaIminjaTabeli,
                                null,
                                null,
                                null,
                                Properties.Settings.Default.dodajZaIzvozVoEksel);

                        newFileBuilder.Append(sqlString);
                        Console.WriteLine("Pretvoriv: " + file.Name + " vo VBA kod! ");
                    }

                    File.AppendAllText(di.FullName + "\\" + moduleFileName, newFileBuilder.ToString());

                    FileInfo[] vbaModuliFajlovi = di.GetFiles("*.bas").OrderBy(fi => fi.Name).ToArray();

                    string makroSodrzhina = File.ReadAllText(vbaModuliFajlovi[0].FullName);
                    string[] makroLinii = makroSodrzhina.Split('\n');

                    string modulSodrzhina = File.ReadAllText(vbaModuliFajlovi[1].FullName);
                    string[] modulLinii = modulSodrzhina.Split('\n');

                    // ostanati linii na makroto
                    StringBuilder writeLines = new StringBuilder();

                    // TODO da se dovrshi ova, pechati po povekje pati edno isto
                    foreach(string makroLinija in makroLinii)
                    {
                        //Console.WriteLine(l);

                        if (!makroLinija.Contains("OpenQuery"))
                        {
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

                                        foreach (string imeTabelaZamena in zamenaIminjaTabeli)
                                        {
                                            del2 = del2.Replace(imeTabelaZamena, PRETSTAVKA_IME_TABELA_KONSTANTA + imeTabelaZamena.ToUpper());
                                        }

                                        metodPotpis = del1 + del2;

                                        string modulIMetoda = fileName + "_module." + metodPotpis;

                                        if (!writeLines.ToString().Contains(modulIMetoda))
                                        {
                                            writeLines.Append("' ----- " + brojRed.ToString() + " ----- " + Environment.NewLine);
                                            writeLines.Append("    Debug.Print " + modulIMetoda + Environment.NewLine);
                                            writeLines.Append("    DoCmd.RunSQL " + modulIMetoda + Environment.NewLine);
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

                    for(int i = 0; i < zamenaIminjaTabeli.Count; i += 1)
                    {
                        if (daliEMegjuTabelaLista[i] && newFileBuilder.ToString().Contains(zamenaIminjaTabeli[i]))
                            dropTablesSB.Append("    DoCmd.RunSQL \"drop table \" & " + PRETSTAVKA_IME_TABELA_KONSTANTA + zamenaIminjaTabeli[i].ToUpper() + Environment.NewLine);
                    }

                    dropTablesSB.Append("End Sub" + Environment.NewLine);

                    File.AppendAllText(di.FullName + "\\" + moduleFileName, dropTablesSB.ToString());

                    Console.WriteLine("Zavrshiv so pretvaranje na kverijata vo VBA kod, fajl: " + moduleFileName);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("IzvrshiPretvoranjeMakroaIKverija(): " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : string.Empty));
            }

            Console.ReadLine();
        }
        
        /// <summary>
        /// Pretvora lista na kverija koi se vo tekstualen format vo dadena papka, vo VBA modul.
        /// Ako ima 100 kverija, kje gi pretvori vo eden VBA modul so 100 funkcii koi vrakjaat SQL string 
        /// (dinamichki SQL). 
        /// </summary>
        public static void IzvrshiPretvoranje()
        {
            try
            {
                // Primeri pr = new Primeri();

                DirectoryInfo sqlKverijaDir = new DirectoryInfo(Properties.Settings.Default.SQL_KVERIJA_PAPKA);
                FileInfo[] sqlKverijaFajlovi = sqlKverijaDir.GetFiles(Properties.Settings.Default.VID_NA_FAJL);
                StringBuilder newFileBuilder = new StringBuilder();
                string fileName = Properties.Settings.Default.IME_NA_IZVEZEN_FAJL;

                DirectoryInfo vbaMakroaKverijaDir = new DirectoryInfo(Properties.Settings.Default.VBA_MAKROA_PAPKA);
                FileInfo[] vbaMakroaKverijaFajlovi = vbaMakroaKverijaDir.GetFiles(Properties.Settings.Default.VBA_MAKRO_VID_FAJL);

                List<string> iminijaTabeli = new List<string>();
                List<string> zamenaIminjaTabeli = new List<string>();

                foreach (string tbl in Properties.Settings.Default.iminijaTabeli)
                {
                    string[] parTabeli = tbl.Split(PretvoriJetSQLKverijaVoVBAUtils.ODDELUVACH_ZA_IMINJA_NA_TABELI);
                    iminijaTabeli.Add(parTabeli[0]);
                    zamenaIminjaTabeli.Add(parTabeli[1]);
                }

                foreach (FileInfo file in sqlKverijaFajlovi)
                {
                    string contents = File.ReadAllText(file.FullName);

                    string sqlString =
                        PretvoriJetSQLKverijaVoVBAUtils.Pretvori(
                            file.Name.Replace(".", "_"),
                            contents,
                            iminijaTabeli,
                            zamenaIminjaTabeli,
                            null,
                            null,
                            null,
                            Properties.Settings.Default.dodajZaIzvozVoEksel);

                    newFileBuilder.Append(sqlString);
                    Console.WriteLine("Pretvoriv: " + file.Name + " vo VBA kod! ");
                }

                File.WriteAllText(Properties.Settings.Default.SQL_KVERIJA_PAPKA + "\\" + fileName, newFileBuilder.ToString());

                Console.WriteLine("Zavrshiv so pretvaranje na kverijata vo VBA kod! ");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Greshka: " + ex.Message);
            }
        }

        public static string ZameniCelZbor(string input, string find, string replace, bool matchWholeWord)
        {
            string textToFind = matchWholeWord ? string.Format(REGEX_TOCHNO_POKLOPUVANJE, find) : find;
            return Regex.Replace(input, textToFind, replace);
        }

        public static string Pretvori(
            string methodName, 
            string kveri, 
            List<string> iminijaTabeli, 
            List<string> zamenaIminjaTabeli,
            string patekaEksel = null,
            string imeIzlezenEkselFajl = null,
            string imeRabotenList = null,
            bool dodajZaIzvozVoEksel = false)
        {
            string tab = "  ";
            StringBuilder ishod = new StringBuilder();
            string[] rows;
            int i = 0;
            int k = 0;
            StringBuilder redMetoda = new StringBuilder();

            if (patekaEksel == null || patekaEksel.Equals(string.Empty))
                patekaEksel = Environment.CurrentDirectory;

            if (imeIzlezenEkselFajl == null || imeIzlezenEkselFajl.Equals(string.Empty))
                imeIzlezenEkselFajl = "imeIzlezenEkselFajl";

            if (imeRabotenList == null || imeRabotenList.Equals(string.Empty))
                imeRabotenList = "imeRabotenList";

            if (methodName.Equals(String.Empty))
                return "methodName e prazen string";

            if (methodName.IndexOf(" ") != -1)
                methodName = methodName.Replace(" ", "_");

            if (methodName.IndexOf("~") != -1)
                methodName = methodName.Replace("~", "");

            if (kveri.Equals(String.Empty))
                return "prazno kveri";

            if (iminijaTabeli.Count == 0)
                return "iminijaTabeli e nula";

            if (zamenaIminjaTabeli.Count == 0)
                return "zamenaIminjaTabeli e nula";

            if (iminijaTabeli.Count != zamenaIminjaTabeli.Count)
                return "brojot na iminja na tabeli i nivnite zameni ne e ednakov: " + iminijaTabeli.Count + ", " + zamenaIminjaTabeli.Count;

            kveri = kveri.Replace("\"", "\"\"");

            foreach (string imeTabela in iminijaTabeli)
            {
                kveri = ZameniCelZbor(kveri, imeTabela, "\" & " + zamenaIminjaTabeli[i] + " & \"", true);
                //kveri = kveri.Replace(imeTabela, "\" & " + zamenaIminjaTabeli[i] + " & \"");
                i += 1;
            }

            imeRabotenList = methodName.Replace("_sql", "");

            if (dodajZaIzvozVoEksel)
                redMetoda.Append("Public Function " + methodName + "(ByVal " + imeRabotenList + " As String, ");
            else
                redMetoda.Append("Public Function " + methodName + "(");

            foreach (string zamena in zamenaIminjaTabeli)
            {
                if (kveri.Contains(zamena))
                {
                    if (Regex.Matches(redMetoda.ToString(), "OU=" + zamena).Count == 0)
                    {
                        redMetoda.Append("ByVal " + zamena + " As String ");

                        if (k >= 0 && k < zamenaIminjaTabeli.Count - 1)
                        {
                            redMetoda.Append(", ");
                        }

                        k += 1;
                    }
                }
            }

            // otstrani posledna zapirka ako metodot ima samo eden parametar. 
            if (redMetoda[redMetoda.Length - 2].Equals(','))
                redMetoda = redMetoda.Remove(redMetoda.Length - 3, 2);

            redMetoda.Append(") As String\n");

            ishod.Append(redMetoda.ToString());
            ishod.Append(tab + "Dim sql as String\n");

            rows = kveri.Split('\n');
            foreach (string row in rows)
            {
                if (row.StartsWith("SELECT"))
                {
                    string[] sqlKoloni = row.Replace(", ", SPEC_KARAKTER_ZA_ZAMENA.ToString()).Split(SPEC_KARAKTER_ZA_ZAMENA);

                    if (sqlKoloni.Length == 1)
                        ishod.Append(tab + "sql = sql & \"" + sqlKoloni[0].Replace("\r", "") + " \" & vbNewLine\n");
                    else
                    {
                        for (int kol = 0; kol < sqlKoloni.Length - 1; kol += 1)
                        {
                            ishod.Append(tab + "sql = sql & \"" + sqlKoloni[kol].Replace("\r", "") + ", \" & vbNewLine\n");
                        }

                        ishod.Append(tab + "sql = sql & \"" + sqlKoloni[sqlKoloni.Length - 1].Replace("\r", "") + " \" & vbNewLine\n");
                    }
                }
                else
                {
                    if (!kveri.ToUpper().Contains("INTO") && row.ToUpper().Contains("FROM"))
                    {
                        if (!Properties.Settings.Default.PATEKA_IZVEZEN_EKSEL_FAJl.Equals(null) && !Properties.Settings.Default.PATEKA_IZVEZEN_EKSEL_FAJl.Equals(string.Empty))
                            patekaEksel = Properties.Settings.Default.PATEKA_IZVEZEN_EKSEL_FAJl.Replace("\\\\", "\\");

                        if (!patekaEksel.EndsWith("\\"))
                            patekaEksel += "\\";

                        if (!Properties.Settings.Default.IME_NA_IZVEZEN_EKSEL_FAJL.Equals(null) && !Properties.Settings.Default.IME_NA_IZVEZEN_EKSEL_FAJL.Equals(string.Empty))
                            imeIzlezenEkselFajl = Properties.Settings.Default.IME_NA_IZVEZEN_EKSEL_FAJL;

                        if (!imeIzlezenEkselFajl.EndsWith(".xls"))
                            imeIzlezenEkselFajl += ".xls";

                        ishod.Append(tab + "sql = sql & \"INTO [\" & " + imeRabotenList + " & \"] IN ''[Excel 8.0;Database=" + patekaEksel + imeIzlezenEkselFajl + "] \"\n");
                    }

                    ishod.Append(tab + "sql = sql & \"" + row.Replace("\r", "") + " \" & vbNewLine\n");
                }
            }

            ishod.Append(tab + methodName + " = sql\n");
            ishod.Append("End Function\n\n");

            return ishod.ToString();
        }
    }
}
