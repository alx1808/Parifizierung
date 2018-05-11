using System;
namespace InterfacesPari
{
    public interface IKategorieRecord
    {
        string Begrundung { get; set; }
        int KategorieID { get; set; }
        string Lage { get; set; }
        double Nutzwert { get; set; }
        int ProjektId { get; set; }
        string RNW { get; set; }
        string Top { get; set; }
        string Widmung { get; set; }
        string Identification { get; }
    }
}
