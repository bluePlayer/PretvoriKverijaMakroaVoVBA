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

        public List<ImaTabelaZamena> iminjaTabeliZameni { get; set; }

        public AppState()
        {
            iminjaTabeliZameni = new List<ImaTabelaZamena>();

            VchitajIminjaTabeliZameni();
        }

        public void VchitajIminjaTabeliZameni()
        {
            foreach (string tbl in Properties.Settings.Default.iminijaTabeli)
            {
                string[] parTabeli = tbl.Split(Konstanti.ODDELUVACH_ZA_IMINJA_NA_TABELI);

                ImaTabelaZamena imaTabelaZamena = new ImaTabelaZamena();
                imaTabelaZamena.imeTabela = parTabeli[0];
                imaTabelaZamena.imeTabelaZamena = parTabeli[1];
                imaTabelaZamena.daliEMegjuTabela = parTabeli[2] == "1";

                iminjaTabeliZameni.Add(imaTabelaZamena);
            }
        }

        // TODO
        public override string ToString()
        {
            return base.ToString();
        }
    }
}
