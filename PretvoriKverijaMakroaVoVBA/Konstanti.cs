using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PretvoriKverijaMakroaVoVBA
{
    public class Konstanti
    {
        public const char SPEC_KARAKTER_ZA_ZAMENA = '@';
        public const char ODDELUVACH_ZA_IMINJA_NA_TABELI = '|';
        public const string PRETSTAVKA_IME_TABELA_KONSTANTA = "TBL_";
        public const string REGEX_TOCHNO_POKLOPUVANJE = @"\b{0}\b";

        public const int SELECT_KVERI = 1;
        public const int SELECT_INTO_KVERI = 2;
        public const int INSET_INTO_KVERI = 3;
        public const int UPDATE_KVERI = 4;
        public const int DELETE_KVERI = 5;
        public const int TRANSFORM_KVERI = 6;
    }
}
