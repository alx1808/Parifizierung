using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InterfacesPari
{
    public interface IZuAbschlagVorgabeRecord
    {
        string Beschreibung { get; set; }
        double Prozent { get; set;  }
    }
}
