using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PretvoriKverijaMakroaVoVBA
{
    class Program
    {
        private static AppState appState;

        static void Main(string[] args)
        {
            appState = new AppState();
            Console.WriteLine(appState);

            PretvoriJetSQLKverijaVoVBAUtils.IzveziIminjaTabeliKonstantiVoFajl();
            PretvoriJetSQLKverijaVoVBAUtils.IzvrshiPretvoranje();
            PretvoriJetSQLKverijaVoVBAUtils.IzvrshiPretvoranjeMakroaIKverija();
            PretvoriJetSQLKverijaVoVBAUtils.DodajTestTabeli();
        }
    }
}
