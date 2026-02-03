using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PretvoriKverijaMakroaVoVBA
{
    public class VbaMakro
    {
        public AppState appState { get; set; }

        public string ime { get; set; }
        public string patekaFajl { get; set; }

        public List<VbaKveri> kverija { get; set; }

        public VbaMakro(AppState appState)
        {
            this.appState = appState;

            kverija = new List<VbaKveri>();
        }

        // TODO
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("ime: " + ime + Environment.NewLine);
            sb.Append("patekaFajl: " + patekaFajl + Environment.NewLine);

            foreach (VbaKveri vbaKveri in kverija)
                sb.Append(vbaKveri.ToString());

            return base.ToString();
        }
    }
}
