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
            foreach (string tbl in Properties.Settings.Default.iminijaTabeli)
            {
                string[] parTabeli = tbl.Split(Konstanti.ODDELUVACH_ZA_IMINJA_NA_TABELI);

                ImeTabelaZamena imaTabelaZamena = new ImeTabelaZamena();
                imaTabelaZamena.ime = parTabeli[0];
                imaTabelaZamena.imeZamena = parTabeli[1];
                imaTabelaZamena.daliEMegjuTabela = parTabeli[2] == "1";

                iminjaTabeliZameni.Add(imaTabelaZamena);
            }
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
