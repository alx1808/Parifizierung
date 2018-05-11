using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UiPari.ViewModel
{
    public class ZuAbschlagVorgabe : IZuAbschlagVorgabeRecord
    {
        private string _Beschreibung = "";
        public string Beschreibung
        {
            get { return _Beschreibung; }
            set { _Beschreibung = value; }
        }
        public double Prozent { get; set; }
    }
}
