using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PretvoriKverijaMakroaVoVBA
{
    public class VbaConvertedMacro
    {
        public AppState appState { get; set; }

        public string ime { get; set; }
        public string patekaFajl { get; set; }

        private string _convertedMacro;
        public string ConvertedMacro 
        {
            get
            {
                return _convertedMacro;
            }
            set
            {
                _convertedMacro = value;
                ConvertedMacroLinii = _convertedMacro.Split('\n');
            }
        }

        public string[] ConvertedMacroLinii { get; set; }

        public VbaConvertedMacro(AppState appState)
        {
            this.appState = appState;
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("----- VbaConvertedMacro -----" + Environment.NewLine);

            sb.Append("ime: " + ime + Environment.NewLine);
            sb.Append("patekaFajl: " + patekaFajl + Environment.NewLine);
            sb.Append("ConvertedMacro: " + ConvertedMacro + Environment.NewLine);

            foreach (string linija in ConvertedMacroLinii)
                sb.Append(linija);

            sb.Append(Environment.NewLine);

            return sb.ToString();
        }
    }
}
