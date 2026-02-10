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

        public List<VbaMakro> listaVbaConvertedMacro { get; set; }
        public List<VbaMakro> listaVbaMakroa { get; set; }

        // TODO da se dodade nova klasa za izvezenite Makroa i tuka lista od istite. 
        // Se misli na makroata pretvoreni od Access vo VBA fajl.

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

            VchitajListaConvertedMacro();
            //VchitajMakroaKverija();

            DodajTestTabeli();
        }

        public void VchitajIminjaTabeliZameni()
        {
            // TODO da se dovrshi
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

        public void VchitajListaConvertedMacro()
        {
            // TODO da se dovrshi
            listaVbaConvertedMacro = new List<VbaMakro>();

            try
            {
                DirectoryInfo sqlMakroaDir = new DirectoryInfo(vbaMakroaPapka);
                DirectoryInfo[] sqlKverijaDir = sqlMakroaDir.GetDirectories();

                foreach (DirectoryInfo di in sqlKverijaDir)
                {
                    FileInfo[] vbaModuliFajlovi = di.GetFiles("*.bas").OrderBy(fi => fi.Name).ToArray();

                    VbaMakro vbaConvertedMacro = new VbaMakro(this, vbaModuliFajlovi[0].Name, vbaModuliFajlovi[0].DirectoryName);

                    listaVbaConvertedMacro.Add(vbaConvertedMacro);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("VchitajListaConvertedMacro(): " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : string.Empty));
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

            File.WriteAllText(sqlKverijaPapka + "\\" + imeNaIzvezenFajl, ishod.ToString());

            Console.WriteLine("Zavrshiv so VBA kod za test tabeli! ");
            Console.ReadLine();
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("----- AppState -----" + Environment.NewLine);

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

            foreach (VbaMakro vbaConvertedMacro in listaVbaConvertedMacro)
                sb.Append(vbaConvertedMacro.ToString());

            //foreach (VbaMakro vbaMakro in listaVbaMakroa)
            //    sb.Append(vbaMakro.ToString());

            sb.Append(Environment.NewLine);

            return sb.ToString();
        }
    }
}
