using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PretvoriKverijaMakroaVoVBA
{
    public class VbaMakro
    {
        public AppState appState { get; set; }

        public string imeIzvodMakro { get; private set; }

        private string _ime;
        public string ime
        {
            get
            {
                return _ime;
            }
            private set
            {
                _ime = value;
                imeIzvodMakro = _ime + "_SQL.bas";
            }
        }

        private string _imeFajl;
        public string imeFajl
        {
            get
            {
                return _imeFajl;
            }
            set
            {
                _imeFajl = value;
                ime = _imeFajl.Split('.')[0];
            }
        }
        public string patekaFajl { get; set; }

        public List<VbaKveri> kverija { get; set; }
        public string[] Linii { get; set; }

        public StringBuilder writeLinesSB;
        public StringBuilder dropTablesSB;

        public DirectoryInfo directoryInfo { get; set; }
        public FileInfo[] fajloviVoPapka { get; set; }

        private string _sodrzhina;
        public string Sodrzhina
        {
            get
            {
                return _sodrzhina;
            }
            set
            {
                _sodrzhina = value;
                Linii = _sodrzhina.Split('\n');
            }
        }

        public VbaMakro(AppState appState, string imeMakro, string pateka)
        {
            this.appState = appState;
            writeLinesSB = new StringBuilder();
            dropTablesSB = new StringBuilder();
            imeFajl = imeMakro;
            patekaFajl = pateka;

            Sodrzhina = File.ReadAllText(patekaFajl + "\\" + imeFajl);

            writeLinesSB.Append("Option Compare Database" + Environment.NewLine + Environment.NewLine);

            directoryInfo = new DirectoryInfo(patekaFajl);
            fajloviVoPapka = directoryInfo.GetFiles(appState.vidNaFajl);

            // gi pretvora kverijata vo funkcii
            foreach (FileInfo file in fajloviVoPapka)
            {
                string contents = File.ReadAllText(file.FullName);

                string sqlString =
                    PretvoriJetSQLKverijaVoVBAUtils.Pretvori(
                        file.Name.Replace(".", "_"),
                        contents,
                        appState.iminjaTabeliZameni,
                        null,
                        null,
                        null);

                writeLinesSB.Append(sqlString);
                Console.WriteLine("Pretvoriv: " + file.Name + " vo VBA kod! ");
            }

            kverija = new List<VbaKveri>();

            foreach (FileInfo file in fajloviVoPapka)
            {
                VbaKveri vbaKveri = new VbaKveri(appState);
                vbaKveri.imeFajl = file.Name;
                vbaKveri.patekaFajl = file.DirectoryName;
                vbaKveri.Kveri = File.ReadAllText(file.FullName);

                kverija.Add(vbaKveri);
            }

            foreach (string makroLinija in Linii)
            {
                if (makroLinija.Contains("End Function"))
                {
                    writeLinesSB.Append("    DoCmd.SetWarnings True" + Environment.NewLine);
                }

                if (!makroLinija.Contains("OpenQuery"))
                {
                    if (!makroLinija.Contains("Option Compare Database"))
                        writeLinesSB.Append(makroLinija + Environment.NewLine);
                }

                if (makroLinija.Contains("SetWarnings"))
                {
                    writeLinesSB.Append("    Dim sql as String" + Environment.NewLine);

                    for (int k = 0; k < kverija.Count; k += 1)
                    {
                        writeLinesSB.Append("    '----- " + (k + 1).ToString() + " ----- " + Environment.NewLine);
                        writeLinesSB.Append("    sql = " + kverija[k].imeFajl.Replace(".", "_") + "(");

                        for (int i = 0; i < kverija[k].tabeli.Count; i += 1)
                        {
                            writeLinesSB.Append(
                                appState.imeFajlTabeliKonstanti.Replace("bas", string.Empty) +
                                Konstanti.PRETSTAVKA_IME_TABELA_KONSTANTA +
                                kverija[k].tabeli[i].imeZamena.ToUpper());

                            if (i < kverija[k].tabeli.Count - 1)
                                writeLinesSB.Append(", ");

                        }

                        writeLinesSB.AppendLine(")" + Environment.NewLine);

                        writeLinesSB.Append("    Debug.Print sql" + Environment.NewLine);
                        writeLinesSB.Append("    DoCmd.RunSQL sql" + Environment.NewLine);
                        writeLinesSB.Append(Environment.NewLine);
                        
                    }
                }
            }

            File.WriteAllText(directoryInfo.FullName + "\\" + imeIzvodMakro, writeLinesSB.ToString() + Environment.NewLine);

            //dropTablesSB.Append("Public Sub brishiMegjuTabeli()" + Environment.NewLine);

            //for (int k = 0; k < kverija.Count; k += 1)
            //{
            //    for (int i = 0; i < kverija[k].tabeli.Count; i += 1)
            //    {
            //        if (!dropTablesSB.ToString().Contains(kverija[k].tabeli[i].imeZamena.ToUpper()) && kverija[k].tabeli[i].daliEMegjuTabela && writeLinesSB.ToString().Contains(kverija[k].tabeli[i].imeZamena) )
            //            dropTablesSB.Append("    DoCmd.RunSQL \"drop table \" & " + appState.imeFajlTabeliKonstanti.Replace("bas", string.Empty) + Konstanti.PRETSTAVKA_IME_TABELA_KONSTANTA + kverija[k].tabeli[i].imeZamena.ToUpper() + Environment.NewLine);
            //    }
            //}

            //dropTablesSB.Append("End Sub" + Environment.NewLine);

            //File.AppendAllText(directoryInfo.FullName + "\\" + imeIzvodMakro, dropTablesSB.ToString());

            Console.WriteLine("Zavrshiv so pretvaranje na kverijata vo VBA kod, fajl: " + imeIzvodMakro);
        }

        // TODO
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("----- VbaMakro -----" + Environment.NewLine);

            sb.Append("ime: " + ime + Environment.NewLine);
            sb.Append("imeFajl: " + imeFajl + Environment.NewLine);
            sb.Append("imeIzvodMakro: " + imeIzvodMakro + Environment.NewLine);
            sb.Append("patekaFajl: " + patekaFajl + Environment.NewLine);
            sb.Append("Sodrzhina: " + Sodrzhina + Environment.NewLine);

            foreach (string linija in Linii)
                sb.Append(linija);

            foreach (VbaKveri vbaKveri in kverija)
                sb.Append(vbaKveri.ToString());

            sb.Append(Environment.NewLine);

            return sb.ToString();
        }
    }
}
