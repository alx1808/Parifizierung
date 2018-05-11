using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MdbPari
{
    internal class SimpleZuAbschlagVorgabeRecord : IZuAbschlagVorgabeRecord
    {
        public string Beschreibung {get;set;}
        public double Prozent {get;set;}

        public override string ToString()
        {
            return Beschreibung;
        }
    }
}
