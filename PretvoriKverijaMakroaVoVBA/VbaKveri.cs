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

        public string ime { get; set; }
        public int vid { get; set; }
        public string kveri { get; set; }

        public List<string> tabeli { get; set; }

        public VbaKveri(AppState appState)
        {
            this.appState = appState;

            tabeli = new List<string>();
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("----- VbaKveri -----" + Environment.NewLine);

            sb.Append("ime: " + ime + Environment.NewLine);
            sb.Append("vid: " + vid.ToString() + Environment.NewLine);
            sb.Append("kveri: " + kveri + Environment.NewLine);

            foreach (string tbl in tabeli)
                sb.Append(tbl + Environment.NewLine);

            return sb.ToString();
        }
    }
}
