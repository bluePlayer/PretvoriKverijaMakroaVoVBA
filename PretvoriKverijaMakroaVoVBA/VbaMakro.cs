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
    }
}
