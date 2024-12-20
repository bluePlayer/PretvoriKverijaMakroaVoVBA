﻿using System;
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

                    foreach (string tbl in Properties.Settings.Default.iminijaTabeli)
                    {
                        string[] parTabeli = tbl.Split(PretvoriJetSQLKverijaVoVBAUtils.ODDELUVACH_ZA_IMINJA_NA_TABELI);
                        iminijaTabeli.Add(parTabeli[0]);
                        zamenaIminjaTabeli.Add(parTabeli[1]);
                    }

                    Console.WriteLine("Rabotam vo papka: " + di.Name);

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

                    moduleFileName = fileName + "_module.bas";

                    File.WriteAllText(di.FullName + "\\" + moduleFileName, newFileBuilder.ToString());

                    FileInfo[] vbaModuliFajlovi = di.GetFiles("*.bas").OrderBy(fi => fi.Name).ToArray();

                    string makroSodrzhina = File.ReadAllText(vbaModuliFajlovi[0].FullName);
                    string[] makroLinii = makroSodrzhina.Split('\n');

                    string modulSodrzhina = File.ReadAllText(vbaModuliFajlovi[1].FullName);
                    string[] modulLinii = modulSodrzhina.Split('\n');

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
                            sqlKverijaFajlovi = di.GetFiles(Properties.Settings.Default.VID_NA_FAJL);

                            foreach (FileInfo file in sqlKverijaFajlovi)
                            {
                                foreach (string modulLinija in modulLinii)
                                {
                                    if (modulLinija.Contains("Public Function " + file.Name.Replace(".", "_")))
                                    {
                                        writeLines.Append(modulLinija + Environment.NewLine);
                                    }
                                }
                            }
                        }
                    }

                    Console.WriteLine(writeLines.ToString());

                    Console.WriteLine("Zavrshiv so pretvaranje na kverijata vo VBA kod, fajl: " + moduleFileName);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("IzvrshiPretvoranjeMakroaIKverija(): " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : string.Empty));
            }

            Console.ReadLine();
        }

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
            string pattern = @"\btest\b";
            string replace = "text";

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
                pattern = @"\b" + imeTabela + "\b";
                //kveri = Regex.Replace(kveri, pattern, "\" & " + zamenaIminjaTabeli[i] + " & \"");
                kveri = kveri.Replace(imeTabela, "\" & " + zamenaIminjaTabeli[i] + " & \"");
                i += 1;
            }

            imeRabotenList = methodName.Replace("_sql", "");

            if (dodajZaIzvozVoEksel)
                ishod.Append("Public Function " + methodName + "(ByVal " + imeRabotenList + " As String, ");
            else
                ishod.Append("Public Function " + methodName + "(");

            foreach (string zamena in zamenaIminjaTabeli)
            {
                if (kveri.Contains(zamena))
                {
                    ishod.Append("ByVal " + zamena + " As String ");

                    if (k >= 0 && k < zamenaIminjaTabeli.Count - 1)
                    {
                        ishod.Append(", ");
                    }

                    k += 1;
                }
            }

            // otstrani posledna zapirka ako metodot ima samo eden parametar. 
            if (ishod[ishod.Length - 2].Equals(','))
                ishod = ishod.Remove(ishod.Length - 3, 2);

            ishod.Append(") As String\n");
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
                    if (dodajZaIzvozVoEksel && row.ToUpper().Contains("FROM"))
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

                    ishod.Append(tab + "sql = sql & \"" + row.Replace("\r", "") + "\" & vbNewLine\n");
                }
            }

            ishod.Append(tab + methodName + " = sql\n");
            ishod.Append("End Function\n\n");

            return ishod.ToString();
        }
    }
}
