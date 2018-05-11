using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InterfacesPari
{
    public interface IRaumZaRecord
    {
        int RaumId { get; set; }
        string AcadHandle { get; set; }
        string Begrundung { get; set; }
        double Flaeche { get; set; }
        IKategorieZaRecord Kategorie { get; set; }
        int KategorieId { get; set; }
        string KatIdentification { get; }
        string Lage { get; set; }
        double Nutzwert { get; set; }
        int ProjektId { get; set; }
        string Raum { get; set; }
        string RNW { get; set; }
        string Top { get; set; }
        string Widmung { get; }
    }
}
