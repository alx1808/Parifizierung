using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MdbPari
{
    internal class SimpleZuAbschlagRecord : IZuAbschlagRecord
    {
        public string Beschreibung { get; set; }
        public double Prozent { get; set; }
        public int KategorieId { get; set; }
        public int ProjektId { get; set; }
        public int ZuAbschlagId { get; set; }
    }
}
