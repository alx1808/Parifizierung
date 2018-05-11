using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AcadPari
{

    public class BlockInfo : IBlockInfo
    {
        public string Raum { get; set; }
        public string Flaeche { get; set; }
        public string Zusatz { get; set; }
        public string Top { get; set; }
        public string Geschoss { get; set; }
        public string Nutzwert { get; set; }
        public string Begrundung { get; set; }
        public string Handle { get; set; }
        public string Widmung { get; set; }
    }
}
