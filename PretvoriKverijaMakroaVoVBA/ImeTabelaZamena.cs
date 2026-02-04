using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PretvoriKverijaMakroaVoVBA
{
    public class ImeTabelaZamena
    {
        public string ime { get; set; }
        public string imeZamena { get; set; }
        public bool daliEMegjuTabela { get; set; }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("----- ImeTabelaZamena -----" + Environment.NewLine);

            sb.Append("ime: " + ime + ", imeZamena: " + imeZamena + ", daliEMegjuTabela: " + daliEMegjuTabela + Environment.NewLine);

            return sb.ToString();
        }
    }
}
