using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MdbPari
{
    public class SimpleWohnungRecord : IWohnungRecord
    {
        public int WohnungId { get; set; }
        public string Top { get; set; }
        public string Typ { get; set; }
        public int ProjektId { get; set; }
        public string Widmung { get; set; }
        public string Nutzwert { get; set; }
    }
}
