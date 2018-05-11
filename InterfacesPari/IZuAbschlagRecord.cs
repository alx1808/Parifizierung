using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InterfacesPari
{
    public interface IZuAbschlagRecord
    {
        string Beschreibung { get; set; }
        double Prozent { get; set; }
        int KategorieId { get; set; }
        int ProjektId { get; set; }
        int ZuAbschlagId { get; set; }
    }
}
