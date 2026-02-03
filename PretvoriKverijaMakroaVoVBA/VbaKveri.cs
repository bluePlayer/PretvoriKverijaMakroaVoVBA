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
    }
}
