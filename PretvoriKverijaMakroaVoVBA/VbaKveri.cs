using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PretvoriKverijaMakroaVoVBA
{
    public class VbaKveri
    {
        public AppState appState { get; set; }

        public string ime { get; private set; }

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

        public int vid { get; set; }

        private string _kveri;
        public string Kveri
        {
            get
            {
                return _kveri;
            }
            set
            {
                _kveri = value;

                tabeli = new List<ImeTabelaZamena>();

                foreach (ImeTabelaZamena tabela in appState.iminjaTabeliZameni)
                {
                    if (_kveri.Contains(tabela.ime))
                        tabeli.Add(tabela);
                }
            }
        }
        public string patekaFajl { get; set; }

        public List<ImeTabelaZamena> tabeli { get; set; }

        public VbaKveri(AppState appState)
        {
            this.appState = appState;
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("----- VbaKveri -----" + Environment.NewLine);

            sb.Append("ime: " + ime + Environment.NewLine);
            sb.Append("vid: " + vid.ToString() + Environment.NewLine);
            sb.Append("kveri: " + Kveri + Environment.NewLine);

            foreach (ImeTabelaZamena tbl in tabeli)
                sb.Append(tbl.ToString() + Environment.NewLine);

            return sb.ToString();
        }
    }
}
