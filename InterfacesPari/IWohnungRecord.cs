using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InterfacesPari
{
    public interface IWohnungRecord
    {
        int WohnungId { get; set;  }
        string Top { get; set; }
        string Typ { get; set; }
        string Widmung { get; set; }
        string Nutzwert { get; set; }
        int ProjektId { get; set; }
    }
}
