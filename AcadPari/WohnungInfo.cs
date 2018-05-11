using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AcadPari
{
    public class WohnungInfo : IWohnungInfo
    {
        public string Top { get; set;}
        public string Typ { get; set;}
        public string Widmung { get; set; }
        public string Nutzwert { get; set; }
    }
}
