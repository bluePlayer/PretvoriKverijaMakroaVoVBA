using System;
using System.Collections.Generic;
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
    }
}
